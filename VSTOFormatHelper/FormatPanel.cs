using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Text;

namespace FormatHelper
{
    public class FormatPanel : Form
    {
        private static readonly string[] Align3Names = { "左对齐", "居中", "右对齐" };
        private static readonly string[] Align4Names = { "左对齐", "居中", "右对齐", "两端对齐" };
        private static readonly string[] BorderStyleNames = { "单线", "虚线", "点线", "点划线" };
        private static readonly string[] AutoFitNames = { "根据内容", "根据窗口", "固定宽度" };

        private static string SafeAlignName(string[] names, int idx)
        {
            return idx >= 0 && idx < names.Length ? names[idx] : names[0];
        }

        private static string SafeBorderStyle(int idx)
        {
            return idx >= 0 && idx < BorderStyleNames.Length ? BorderStyleNames[idx] : BorderStyleNames[0];
        }

        private static string SafeAutoFit(int idx)
        {
            return idx >= 0 && idx < AutoFitNames.Length ? AutoFitNames[idx] : AutoFitNames[0];
        }

        private TabControl tabControl;
        private ComboBox presetCombo;
        private FormatPreset currentPreset;

        private ComboBox txtChineseFont, txtEnglishFont, txtFontSize;
        private CheckBox txtBold, txtItalic;
        private ComboBox txtLineSpacing, txtSpaceBefore, txtSpaceAfter;

        private ComboBox h1Font, h1Size, h2Font, h2Size, h3Font, h3Size, h4Font, h4Size;
        private CheckBox h1Bold, h2Bold, h3Bold, h4Bold;

        private ComboBox paraAlignment, paraFirstLineIndent, paraLeftIndent, paraRightIndent;
        private ComboBox paraSpaceBefore, paraSpaceAfter;

        private CheckBox tblTopShow, tblBottomShow, tblLeftShow, tblRightShow, tblInsideHShow, tblInsideVShow;
        private ComboBox tblTopStyle, tblTopWidth, tblBottomStyle, tblBottomWidth;
        private ComboBox tblLeftStyle, tblLeftWidth, tblRightStyle, tblRightWidth;
        private ComboBox tblInsideHStyle, tblInsideHWidth, tblInsideVStyle, tblInsideVWidth;
        private ComboBox tblAutoFit, tblFixedWidth;
        private ComboBox tblHeaderFont, tblHeaderSize; private CheckBox tblHeaderBold;
        private ComboBox tblBodyFont, tblBodySize;
        private ComboBox tblNumberFont, tblNumberSize, tblNumberAlignment;
        private CheckBox tblUseThousandSep; private ComboBox tblDecimalPlaces;

        private ComboBox imgMaxWidth, imgMaxHeight; private CheckBox imgKeepRatio; private ComboBox imgAlignment;

        private ComboBox numFont, numSize, numAlignment;
        private CheckBox numUseThousandSep; private ComboBox numDecimalPlaces;

        public FormatPanel(FormatPreset preset)
        {
            currentPreset = preset;
            InitializeComponents();
            LoadPresetToUI(preset);
        }

        private ComboBox CreateFontCombo()
        {
            ComboBox cb = new ComboBox();
            cb.DropDownStyle = ComboBoxStyle.DropDownList;
            cb.Items.Add("宋体");
            cb.Items.Add("黑体");
            cb.Items.Add("楷体");
            cb.Items.Add("仿宋");
            cb.Items.Add("微软雅黑");
            cb.Items.Add("Times New Roman");
            cb.Items.Add("Arial");
            cb.Items.Add("Calibri");
            cb.Items.Add("方正小标宋简体");
            try
            {
                using (InstalledFontCollection fc = new InstalledFontCollection())
                {
                    foreach (System.Drawing.FontFamily ff in fc.Families)
                    {
                        if (!cb.Items.Contains(ff.Name))
                            cb.Items.Add(ff.Name);
                    }
                }
            }
            catch { }
            cb.Width = 140;
            return cb;
        }

        private ComboBox CreateSizeCombo(float[] sizes, float defaultVal)
        {
            ComboBox cb = new ComboBox();
            cb.DropDownStyle = ComboBoxStyle.DropDownList;
            bool hasDefault = false;
            foreach (float s in sizes)
            {
                cb.Items.Add(s.ToString());
                if (Math.Abs(s - defaultVal) < 0.01f) hasDefault = true;
            }
            if (!hasDefault) cb.Items.Add(defaultVal.ToString());
            cb.Text = defaultVal.ToString();
            cb.Width = 70;
            return cb;
        }

        private ComboBox CreateSimpleCombo(string[] items, string defaultVal)
        {
            ComboBox cb = new ComboBox();
            cb.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (string s in items) cb.Items.Add(s);
            cb.Text = defaultVal;
            cb.Width = 100;
            return cb;
        }

        private Label L(string text) { return new Label() { Text = text, AutoSize = true }; }

        private void InitializeComponents()
        {
            this.Text = "格式设置";
            this.Size = new Size(620, 580);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 40 };
            topPanel.Controls.Add(new Label() { Text = "预设方案:", Left = 10, Top = 10, AutoSize = true });
            presetCombo = new ComboBox() { Left = 80, Top = 7, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (FormatPreset p in PresetManager.Presets) presetCombo.Items.Add(p.Name);
            presetCombo.SelectedIndex = 0;
            presetCombo.SelectedIndexChanged += PresetCombo_SelectedIndexChanged;
            topPanel.Controls.Add(presetCombo);
            Button btnSavePreset = new Button() { Text = "保存当前预设", Left = 250, Top = 5, Width = 100 };
            btnSavePreset.Click += BtnSavePreset_Click;
            topPanel.Controls.Add(btnSavePreset);
            this.Controls.Add(topPanel);

            tabControl = new TabControl() { Left = 10, Top = 50, Width = 580, Height = 420 };
            tabControl.SizeMode = TabSizeMode.Fixed;
            tabControl.ItemSize = new Size(80, 25);

            CreateTextTab();
            CreateHeadingTab();
            CreateParagraphTab();
            CreateTableTab();
            CreateImageTab();
            CreateNumberTab();

            this.Controls.Add(tabControl);

            Panel bottomPanel = new Panel() { Dock = DockStyle.Bottom, Height = 50 };
            Button btnApply = new Button() { Text = "应用", Left = 150, Top = 10, Width = 80 };
            btnApply.Click += BtnApply_Click;
            Button btnSave = new Button() { Text = "保存预设", Left = 250, Top = 10, Width = 80 };
            btnSave.Click += BtnSavePreset_Click;
            Button btnReset = new Button() { Text = "重置", Left = 350, Top = 10, Width = 80 };
            btnReset.Click += BtnReset_Click;
            bottomPanel.Controls.Add(btnApply);
            bottomPanel.Controls.Add(btnSave);
            bottomPanel.Controls.Add(btnReset);
            this.Controls.Add(bottomPanel);
        }

        private void CreateTextTab()
        {
            TabPage tab = new TabPage("正文与字体");
            int y = 15;
            tab.Controls.Add(L("中文字体:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            txtChineseFont = CreateFontCombo(); txtChineseFont.Left = 120; txtChineseFont.Top = y - 3; tab.Controls.Add(txtChineseFont);

            tab.Controls.Add(L("英文字体:")); tab.Controls[tab.Controls.Count - 1].Left = 300; tab.Controls[tab.Controls.Count - 1].Top = y;
            txtEnglishFont = CreateFontCombo(); txtEnglishFont.Left = 400; txtEnglishFont.Top = y - 3; tab.Controls.Add(txtEnglishFont);
            y += 35;

            tab.Controls.Add(L("字号:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            txtFontSize = CreateSizeCombo(new float[] { 9, 10.5f, 12, 14, 16, 18, 22, 24, 26, 36 }, 12);
            txtFontSize.Left = 120; txtFontSize.Top = y - 3; tab.Controls.Add(txtFontSize);

            txtBold = new CheckBox() { Text = "加粗", Left = 300, Top = y }; tab.Controls.Add(txtBold);
            txtItalic = new CheckBox() { Text = "斜体", Left = 380, Top = y }; tab.Controls.Add(txtItalic);
            y += 35;

            tab.Controls.Add(L("行距:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            txtLineSpacing = CreateSimpleCombo(new string[] { "1", "1.15", "1.5", "2", "2.5", "3" }, "1.5");
            txtLineSpacing.Left = 120; txtLineSpacing.Top = y - 3; tab.Controls.Add(txtLineSpacing);

            tab.Controls.Add(L("段前:")); tab.Controls[tab.Controls.Count - 1].Left = 250; tab.Controls[tab.Controls.Count - 1].Top = y;
            txtSpaceBefore = CreateSimpleCombo(new string[] { "0", "3", "6", "12", "18", "24" }, "0");
            txtSpaceBefore.Left = 310; txtSpaceBefore.Top = y - 3; tab.Controls.Add(txtSpaceBefore);

            tab.Controls.Add(L("段后:")); tab.Controls[tab.Controls.Count - 1].Left = 420; tab.Controls[tab.Controls.Count - 1].Top = y;
            txtSpaceAfter = CreateSimpleCombo(new string[] { "0", "3", "6", "12", "18", "24" }, "0");
            txtSpaceAfter.Left = 480; txtSpaceAfter.Top = y - 3; tab.Controls.Add(txtSpaceAfter);

            tabControl.Controls.Add(tab);
        }

        private void AddHeadingRow(TabPage tab, string label, ref ComboBox fontCombo, ref ComboBox sizeCombo, ref CheckBox boldCheck, ref int y)
        {
            tab.Controls.Add(L(label)); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            fontCombo = CreateFontCombo(); fontCombo.Left = 120; fontCombo.Top = y - 3; tab.Controls.Add(fontCombo);
            sizeCombo = CreateSizeCombo(new float[] { 12, 14, 16, 18, 22, 24, 26, 36 }, 16);
            sizeCombo.Left = 280; sizeCombo.Top = y - 3; tab.Controls.Add(sizeCombo);
            boldCheck = new CheckBox() { Text = "加粗", Left = 370, Top = y }; tab.Controls.Add(boldCheck);
            y += 35;
        }

        private void CreateHeadingTab()
        {
            TabPage tab = new TabPage("标题");
            int y = 15;
            AddHeadingRow(tab, "一级标题:", ref h1Font, ref h1Size, ref h1Bold, ref y);
            AddHeadingRow(tab, "二级标题:", ref h2Font, ref h2Size, ref h2Bold, ref y);
            AddHeadingRow(tab, "三级标题:", ref h3Font, ref h3Size, ref h3Bold, ref y);
            AddHeadingRow(tab, "四级标题:", ref h4Font, ref h4Size, ref h4Bold, ref y);
            tabControl.Controls.Add(tab);
        }

        private void CreateParagraphTab()
        {
            TabPage tab = new TabPage("段落");
            int y = 15;
            tab.Controls.Add(L("对齐方式:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            paraAlignment = CreateSimpleCombo(new string[] { "左对齐", "居中", "右对齐", "两端对齐" }, "左对齐");
            paraAlignment.Left = 120; paraAlignment.Top = y - 3; tab.Controls.Add(paraAlignment);
            y += 35;

            tab.Controls.Add(L("首行缩进:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            paraFirstLineIndent = CreateSimpleCombo(new string[] { "0", "1", "2", "3", "4" }, "2");
            paraFirstLineIndent.Left = 120; paraFirstLineIndent.Top = y - 3; tab.Controls.Add(paraFirstLineIndent);

            tab.Controls.Add(L("左缩进:")); tab.Controls[tab.Controls.Count - 1].Left = 250; tab.Controls[tab.Controls.Count - 1].Top = y;
            paraLeftIndent = CreateSimpleCombo(new string[] { "0", "1", "2", "3", "4" }, "0");
            paraLeftIndent.Left = 330; paraLeftIndent.Top = y - 3; tab.Controls.Add(paraLeftIndent);

            tab.Controls.Add(L("右缩进:")); tab.Controls[tab.Controls.Count - 1].Left = 420; tab.Controls[tab.Controls.Count - 1].Top = y;
            paraRightIndent = CreateSimpleCombo(new string[] { "0", "1", "2", "3", "4" }, "0");
            paraRightIndent.Left = 500; paraRightIndent.Top = y - 3; tab.Controls.Add(paraRightIndent);
            y += 35;

            tab.Controls.Add(L("段前:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            paraSpaceBefore = CreateSimpleCombo(new string[] { "0", "3", "6", "12", "18", "24" }, "0");
            paraSpaceBefore.Left = 120; paraSpaceBefore.Top = y - 3; tab.Controls.Add(paraSpaceBefore);

            tab.Controls.Add(L("段后:")); tab.Controls[tab.Controls.Count - 1].Left = 250; tab.Controls[tab.Controls.Count - 1].Top = y;
            paraSpaceAfter = CreateSimpleCombo(new string[] { "0", "3", "6", "12", "18", "24" }, "0");
            paraSpaceAfter.Left = 330; paraSpaceAfter.Top = y - 3; tab.Controls.Add(paraSpaceAfter);

            tabControl.Controls.Add(tab);
        }

        private void AddBorderRow(TabPage tab, string label, ref CheckBox showCheck, ref ComboBox styleCombo, ref ComboBox widthCombo, ref int y)
        {
            tab.Controls.Add(L(label)); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            showCheck = new CheckBox() { Text = "显示", Left = 80, Top = y, Checked = true }; tab.Controls.Add(showCheck);
            styleCombo = CreateSimpleCombo(new string[] { "单线", "虚线", "点线", "点划线" }, "单线");
            styleCombo.Left = 150; styleCombo.Top = y - 3; styleCombo.Width = 80; tab.Controls.Add(styleCombo);
            widthCombo = CreateSimpleCombo(new string[] { "0.5", "0.75", "1", "1.5", "2", "3" }, "0.5");
            widthCombo.Left = 250; widthCombo.Top = y - 3; widthCombo.Width = 70; tab.Controls.Add(widthCombo);
            y += 28;
        }

        private void CreateTableTab()
        {
            TabPage tab = new TabPage("表格");
            int y = 10;

            tab.Controls.Add(L("边框设置:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            y += 22;
            AddBorderRow(tab, "上边框:", ref tblTopShow, ref tblTopStyle, ref tblTopWidth, ref y);
            AddBorderRow(tab, "下边框:", ref tblBottomShow, ref tblBottomStyle, ref tblBottomWidth, ref y);
            AddBorderRow(tab, "左边框:", ref tblLeftShow, ref tblLeftStyle, ref tblLeftWidth, ref y);
            AddBorderRow(tab, "右边框:", ref tblRightShow, ref tblRightStyle, ref tblRightWidth, ref y);
            AddBorderRow(tab, "内横线:", ref tblInsideHShow, ref tblInsideHStyle, ref tblInsideHWidth, ref y);
            AddBorderRow(tab, "内竖线:", ref tblInsideVShow, ref tblInsideVStyle, ref tblInsideVWidth, ref y);

            y += 5;
            tab.Controls.Add(L("自动调整:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblAutoFit = CreateSimpleCombo(new string[] { "根据内容", "根据窗口", "固定宽度" }, "根据窗口");
            tblAutoFit.Left = 120; tblAutoFit.Top = y - 3; tab.Controls.Add(tblAutoFit);

            tab.Controls.Add(L("固定宽度:")); tab.Controls[tab.Controls.Count - 1].Left = 280; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblFixedWidth = CreateSimpleCombo(new string[] { "10", "12", "14", "16", "18" }, "14");
            tblFixedWidth.Left = 370; tblFixedWidth.Top = y - 3; tblFixedWidth.Width = 70; tab.Controls.Add(tblFixedWidth);
            y += 30;

            tab.Controls.Add(L("表头字体:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblHeaderFont = CreateFontCombo(); tblHeaderFont.Left = 120; tblHeaderFont.Top = y - 3; tab.Controls.Add(tblHeaderFont);
            tblHeaderSize = CreateSizeCombo(new float[] { 9, 10.5f, 12, 14 }, 10.5f);
            tblHeaderSize.Left = 280; tblHeaderSize.Top = y - 3; tab.Controls.Add(tblHeaderSize);
            tblHeaderBold = new CheckBox() { Text = "加粗", Left = 370, Top = y }; tab.Controls.Add(tblHeaderBold);
            y += 30;

            tab.Controls.Add(L("正文字体:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblBodyFont = CreateFontCombo(); tblBodyFont.Left = 120; tblBodyFont.Top = y - 3; tab.Controls.Add(tblBodyFont);
            tblBodySize = CreateSizeCombo(new float[] { 9, 10.5f, 12, 14 }, 10.5f);
            tblBodySize.Left = 280; tblBodySize.Top = y - 3; tab.Controls.Add(tblBodySize);
            y += 30;

            tab.Controls.Add(L("数字字体:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblNumberFont = CreateFontCombo(); tblNumberFont.Left = 120; tblNumberFont.Top = y - 3; tab.Controls.Add(tblNumberFont);
            tblNumberSize = CreateSizeCombo(new float[] { 9, 10.5f, 12, 14 }, 10.5f);
            tblNumberSize.Left = 280; tblNumberSize.Top = y - 3; tab.Controls.Add(tblNumberSize);
            y += 30;

            tblUseThousandSep = new CheckBox() { Text = "千分符", Left = 20, Top = y }; tab.Controls.Add(tblUseThousandSep);
            tab.Controls.Add(L("小数位数:")); tab.Controls[tab.Controls.Count - 1].Left = 120; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblDecimalPlaces = CreateSimpleCombo(new string[] { "0", "1", "2", "3", "4" }, "2");
            tblDecimalPlaces.Left = 200; tblDecimalPlaces.Top = y - 3; tblDecimalPlaces.Width = 60; tab.Controls.Add(tblDecimalPlaces);
            tab.Controls.Add(L("对齐:")); tab.Controls[tab.Controls.Count - 1].Left = 280; tab.Controls[tab.Controls.Count - 1].Top = y;
            tblNumberAlignment = CreateSimpleCombo(new string[] { "左对齐", "居中", "右对齐" }, "左对齐");
            tblNumberAlignment.Left = 330; tblNumberAlignment.Top = y - 3; tab.Controls.Add(tblNumberAlignment);

            tabControl.Controls.Add(tab);
        }

        private void CreateImageTab()
        {
            TabPage tab = new TabPage("图片");
            int y = 15;
            tab.Controls.Add(L("最大宽度(cm):")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            imgMaxWidth = CreateSimpleCombo(new string[] { "10", "12", "14", "16", "18", "20" }, "14");
            imgMaxWidth.Left = 150; imgMaxWidth.Top = y - 3; tab.Controls.Add(imgMaxWidth);
            y += 35;

            tab.Controls.Add(L("最大高度(cm):")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            imgMaxHeight = CreateSimpleCombo(new string[] { "10", "15", "20", "25", "28" }, "20");
            imgMaxHeight.Left = 150; imgMaxHeight.Top = y - 3; tab.Controls.Add(imgMaxHeight);
            y += 35;

            imgKeepRatio = new CheckBox() { Text = "保持纵横比", Left = 20, Top = y, Checked = true }; tab.Controls.Add(imgKeepRatio);
            y += 35;

            tab.Controls.Add(L("对齐方式:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            imgAlignment = CreateSimpleCombo(new string[] { "左对齐", "居中", "右对齐" }, "居中");
            imgAlignment.Left = 120; imgAlignment.Top = y - 3; tab.Controls.Add(imgAlignment);

            tabControl.Controls.Add(tab);
        }

        private void CreateNumberTab()
        {
            TabPage tab = new TabPage("数字");
            int y = 15;
            tab.Controls.Add(L("字体:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            numFont = CreateFontCombo(); numFont.Left = 120; numFont.Top = y - 3; tab.Controls.Add(numFont);
            y += 35;

            tab.Controls.Add(L("字号:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            numSize = CreateSizeCombo(new float[] { 9, 10.5f, 12, 14, 16 }, 12);
            numSize.Left = 120; numSize.Top = y - 3; tab.Controls.Add(numSize);
            y += 35;

            numUseThousandSep = new CheckBox() { Text = "千分符", Left = 20, Top = y, Checked = true }; tab.Controls.Add(numUseThousandSep);
            y += 30;

            tab.Controls.Add(L("小数位数:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            numDecimalPlaces = CreateSimpleCombo(new string[] { "0", "1", "2", "3", "4" }, "2");
            numDecimalPlaces.Left = 120; numDecimalPlaces.Top = y - 3; numDecimalPlaces.Width = 60; tab.Controls.Add(numDecimalPlaces);
            y += 35;

            tab.Controls.Add(L("对齐方式:")); tab.Controls[tab.Controls.Count - 1].Left = 20; tab.Controls[tab.Controls.Count - 1].Top = y;
            numAlignment = CreateSimpleCombo(new string[] { "左对齐", "居中", "右对齐" }, "左对齐");
            numAlignment.Left = 120; numAlignment.Top = y - 3; tab.Controls.Add(numAlignment);

            tabControl.Controls.Add(tab);
        }

        private void PresetCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            FormatPreset p = PresetManager.GetPresetByName(presetCombo.Text);
            if (p != null)
            {
                currentPreset = p;
                LoadPresetToUI(p);
            }
        }

        private void LoadPresetToUI(FormatPreset p)
        {
            SelectComboByText(txtChineseFont, p.Text.ChineseFont);
            SelectComboByText(txtEnglishFont, p.Text.EnglishFont);
            SelectComboByText(txtFontSize, p.Text.FontSize.ToString());
            txtBold.Checked = p.Text.Bold;
            txtItalic.Checked = p.Text.Italic;
            SelectComboByText(txtLineSpacing, p.Text.LineSpacing.ToString());
            SelectComboByText(txtSpaceBefore, p.Text.SpaceBefore.ToString());
            SelectComboByText(txtSpaceAfter, p.Text.SpaceAfter.ToString());

            SelectComboByText(h1Font, p.Heading.Heading1Font);
            SelectComboByText(h1Size, p.Heading.Heading1Size.ToString());
            h1Bold.Checked = p.Heading.Heading1Bold;
            SelectComboByText(h2Font, p.Heading.Heading2Font);
            SelectComboByText(h2Size, p.Heading.Heading2Size.ToString());
            h2Bold.Checked = p.Heading.Heading2Bold;
            SelectComboByText(h3Font, p.Heading.Heading3Font);
            SelectComboByText(h3Size, p.Heading.Heading3Size.ToString());
            h3Bold.Checked = p.Heading.Heading3Bold;
            SelectComboByText(h4Font, p.Heading.Heading4Font);
            SelectComboByText(h4Size, p.Heading.Heading4Size.ToString());
            h4Bold.Checked = p.Heading.Heading4Bold;

            SelectComboByText(paraAlignment, SafeAlignName(Align4Names, p.Paragraph.Alignment));
            SelectComboByText(paraFirstLineIndent, p.Paragraph.FirstLineIndent.ToString());
            SelectComboByText(paraLeftIndent, p.Paragraph.LeftIndent.ToString());
            SelectComboByText(paraRightIndent, p.Paragraph.RightIndent.ToString());
            SelectComboByText(paraSpaceBefore, p.Paragraph.SpaceBefore.ToString());
            SelectComboByText(paraSpaceAfter, p.Paragraph.SpaceAfter.ToString());

            LoadBorderUI(tblTopShow, tblTopStyle, tblTopWidth, p.Table.TopBorder);
            LoadBorderUI(tblBottomShow, tblBottomStyle, tblBottomWidth, p.Table.BottomBorder);
            LoadBorderUI(tblLeftShow, tblLeftStyle, tblLeftWidth, p.Table.LeftBorder);
            LoadBorderUI(tblRightShow, tblRightStyle, tblRightWidth, p.Table.RightBorder);
            LoadBorderUI(tblInsideHShow, tblInsideHStyle, tblInsideHWidth, p.Table.InsideHorizontalBorder);
            LoadBorderUI(tblInsideVShow, tblInsideVStyle, tblInsideVWidth, p.Table.InsideVerticalBorder);

            SelectComboByText(tblAutoFit, SafeAutoFit(p.Table.AutoFit));
            SelectComboByText(tblFixedWidth, p.Table.FixedWidth.ToString());
            SelectComboByText(tblHeaderFont, p.Table.HeaderFont);
            SelectComboByText(tblHeaderSize, p.Table.HeaderFontSize.ToString());
            tblHeaderBold.Checked = p.Table.HeaderBold;
            SelectComboByText(tblBodyFont, p.Table.BodyFont);
            SelectComboByText(tblBodySize, p.Table.BodyFontSize.ToString());
            SelectComboByText(tblNumberFont, p.Table.NumberFont);
            SelectComboByText(tblNumberSize, p.Table.NumberFontSize.ToString());
            tblUseThousandSep.Checked = p.Table.UseThousandSeparator;
            SelectComboByText(tblDecimalPlaces, p.Table.DecimalPlaces.ToString());
            SelectComboByText(tblNumberAlignment, SafeAlignName(Align3Names, p.Table.NumberAlignment));

            SelectComboByText(imgMaxWidth, p.Image.MaxWidth.ToString());
            SelectComboByText(imgMaxHeight, p.Image.MaxHeight.ToString());
            imgKeepRatio.Checked = p.Image.KeepAspectRatio;
            SelectComboByText(imgAlignment, SafeAlignName(Align3Names, p.Image.Alignment));

            SelectComboByText(numFont, p.Number.NumberFont);
            SelectComboByText(numSize, p.Number.NumberFontSize.ToString());
            numUseThousandSep.Checked = p.Number.UseThousandSeparator;
            SelectComboByText(numDecimalPlaces, p.Number.DecimalPlaces.ToString());
            SelectComboByText(numAlignment, SafeAlignName(Align3Names, p.Number.Alignment));
        }

        private void LoadBorderUI(CheckBox showCheck, ComboBox styleCombo, ComboBox widthCombo, BorderDef def)
        {
            showCheck.Checked = def.Show;
            SelectComboByText(styleCombo, SafeBorderStyle(def.Style));
            SelectComboByText(widthCombo, def.Width.ToString());
        }

        private void SelectComboByText(ComboBox cb, string text)
        {
            for (int i = 0; i < cb.Items.Count; i++)
            {
                if (cb.Items[i].ToString() == text) { cb.SelectedIndex = i; return; }
            }
            if (cb.DropDownStyle == ComboBoxStyle.DropDownList)
            {
                cb.Items.Add(text);
                cb.Text = text;
            }
        }

        private FormatPreset SaveUIToPreset()
        {
            FormatPreset p = new FormatPreset();
            p.Name = presetCombo.Text;

            p.Text.ChineseFont = txtChineseFont.Text;
            p.Text.EnglishFont = txtEnglishFont.Text;
            p.Text.FontSize = ParseFloat(txtFontSize.Text, 12f);
            p.Text.Bold = txtBold.Checked;
            p.Text.Italic = txtItalic.Checked;
            p.Text.LineSpacing = ParseFloat(txtLineSpacing.Text, 1.5f);
            p.Text.SpaceBefore = ParseFloat(txtSpaceBefore.Text, 0f);
            p.Text.SpaceAfter = ParseFloat(txtSpaceAfter.Text, 0f);

            p.Heading.Heading1Font = h1Font.Text;
            p.Heading.Heading1Size = ParseFloat(h1Size.Text, 22f);
            p.Heading.Heading1Bold = h1Bold.Checked;
            p.Heading.Heading2Font = h2Font.Text;
            p.Heading.Heading2Size = ParseFloat(h2Size.Text, 16f);
            p.Heading.Heading2Bold = h2Bold.Checked;
            p.Heading.Heading3Font = h3Font.Text;
            p.Heading.Heading3Size = ParseFloat(h3Size.Text, 14f);
            p.Heading.Heading3Bold = h3Bold.Checked;
            p.Heading.Heading4Font = h4Font.Text;
            p.Heading.Heading4Size = ParseFloat(h4Size.Text, 14f);
            p.Heading.Heading4Bold = h4Bold.Checked;

            p.Paragraph.Alignment = paraAlignment.SelectedIndex;
            p.Paragraph.FirstLineIndent = ParseFloat(paraFirstLineIndent.Text, 2f);
            p.Paragraph.LeftIndent = ParseFloat(paraLeftIndent.Text, 0f);
            p.Paragraph.RightIndent = ParseFloat(paraRightIndent.Text, 0f);
            p.Paragraph.SpaceBefore = ParseFloat(paraSpaceBefore.Text, 0f);
            p.Paragraph.SpaceAfter = ParseFloat(paraSpaceAfter.Text, 0f);

            SaveBorderUI(tblTopShow, tblTopStyle, tblTopWidth, p.Table.TopBorder);
            SaveBorderUI(tblBottomShow, tblBottomStyle, tblBottomWidth, p.Table.BottomBorder);
            SaveBorderUI(tblLeftShow, tblLeftStyle, tblLeftWidth, p.Table.LeftBorder);
            SaveBorderUI(tblRightShow, tblRightStyle, tblRightWidth, p.Table.RightBorder);
            SaveBorderUI(tblInsideHShow, tblInsideHStyle, tblInsideHWidth, p.Table.InsideHorizontalBorder);
            SaveBorderUI(tblInsideVShow, tblInsideVStyle, tblInsideVWidth, p.Table.InsideVerticalBorder);

            p.Table.AutoFit = tblAutoFit.SelectedIndex;
            p.Table.FixedWidth = ParseFloat(tblFixedWidth.Text, 14f);
            p.Table.HeaderFont = tblHeaderFont.Text;
            p.Table.HeaderFontSize = ParseFloat(tblHeaderSize.Text, 10.5f);
            p.Table.HeaderBold = tblHeaderBold.Checked;
            p.Table.BodyFont = tblBodyFont.Text;
            p.Table.BodyFontSize = ParseFloat(tblBodySize.Text, 10.5f);
            p.Table.NumberFont = tblNumberFont.Text;
            p.Table.NumberFontSize = ParseFloat(tblNumberSize.Text, 10.5f);
            p.Table.UseThousandSeparator = tblUseThousandSep.Checked;
            p.Table.DecimalPlaces = ParseInt(tblDecimalPlaces.Text, 2);
            p.Table.NumberAlignment = tblNumberAlignment.SelectedIndex;

            p.Image.MaxWidth = ParseFloat(imgMaxWidth.Text, 14f);
            p.Image.MaxHeight = ParseFloat(imgMaxHeight.Text, 20f);
            p.Image.KeepAspectRatio = imgKeepRatio.Checked;
            p.Image.Alignment = imgAlignment.SelectedIndex;

            p.Number.NumberFont = numFont.Text;
            p.Number.NumberFontSize = ParseFloat(numSize.Text, 12f);
            p.Number.UseThousandSeparator = numUseThousandSep.Checked;
            p.Number.DecimalPlaces = ParseInt(numDecimalPlaces.Text, 2);
            p.Number.Alignment = numAlignment.SelectedIndex;

            return p;
        }

        private void SaveBorderUI(CheckBox showCheck, ComboBox styleCombo, ComboBox widthCombo, BorderDef def)
        {
            def.Show = showCheck.Checked;
            def.Style = styleCombo.SelectedIndex;
            def.Width = ParseFloat(widthCombo.Text, 0.5f);
        }

        private float ParseFloat(string s, float def)
        {
            float v;
            if (float.TryParse(s, out v)) return v;
            return def;
        }

        private int ParseInt(string s, int def)
        {
            int v;
            if (int.TryParse(s, out v)) return v;
            return def;
        }

        private void BtnSavePreset_Click(object sender, EventArgs e)
        {
            FormatPreset p = SaveUIToPreset();
            FormatPreset existing = PresetManager.GetPresetByName(p.Name);
            if (existing != null)
            {
                int idx = PresetManager.Presets.IndexOf(existing);
                PresetManager.Presets[idx] = p;
            }
            else
            {
                PresetManager.Presets.Add(p);
            }
            currentPreset = p;
            PresetManager.SavePresets();
            MessageBox.Show("预设 \"" + p.Name + "\" 已保存", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            currentPreset = SaveUIToPreset();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            FormatPreset p = PresetManager.GetPresetByName(presetCombo.Text);
            if (p != null) LoadPresetToUI(p);
        }

        public FormatPreset GetPreset()
        {
            return currentPreset;
        }
    }
}
