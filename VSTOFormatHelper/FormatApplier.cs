using System;
using Microsoft.Office.Interop.Word;

namespace FormatHelper
{
    public class FormatApplier
    {
        private Document _doc;

        public FormatApplier(Document doc)
        {
            _doc = doc;
        }

        public void ApplyAll(FormatPreset preset)
        {
            ApplyHeadingFormat(preset.Heading);
            ApplyTextFormat(preset.Text);
            ApplyParagraphFormat(preset.Paragraph);
            ApplyTableFormat(preset.Table);
            ApplyImageFormat(preset.Image);
            // ApplyNumberFormat is for standalone number paragraphs only;
            // table numbers are handled inside ApplyTableFormat.
            ApplyNumberFormat(preset.Number);
        }

        private bool IsHeadingStyle(Style style)
        {
            try
            {
                string name = style.NameLocal;
                if (name == "标题 1" || name == "标题 2" || name == "标题 3" || name == "标题 4" ||
                    name == "Heading 1" || name == "Heading 2" || name == "Heading 3" || name == "Heading 4")
                    return true;
            }
            catch { }
            return false;
        }

        public void ApplyTextFormat(TextFormat fmt)
        {
            foreach (Paragraph para in _doc.Paragraphs)
            {
                try
                {
                    Style style = (Style)para.get_Style();
                    if (IsHeadingStyle(style)) continue;
                    Range r = para.Range;
                    if ((int)r.get_Information(WdInformation.wdWithInTable) != 0) continue;

                    r.Font.Name = fmt.ChineseFont;
                    r.Font.NameOther = fmt.EnglishFont;
                    r.Font.Size = fmt.FontSize;
                    r.Font.Bold = fmt.Bold ? 1 : 0;
                    r.Font.Italic = fmt.Italic ? 1 : 0;
                    para.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                    para.Format.LineSpacing = fmt.LineSpacing;
                    para.Format.SpaceBefore = fmt.SpaceBefore;
                    para.Format.SpaceAfter = fmt.SpaceAfter;
                }
                catch { }
            }
        }

        public void ApplyHeadingFormat(HeadingFormat fmt)
        {
            ApplySingleHeadingStyle("标题 1", "Heading 1", fmt.Heading1Font, fmt.Heading1Size, fmt.Heading1Bold);
            ApplySingleHeadingStyle("标题 2", "Heading 2", fmt.Heading2Font, fmt.Heading2Size, fmt.Heading2Bold);
            ApplySingleHeadingStyle("标题 3", "Heading 3", fmt.Heading3Font, fmt.Heading3Size, fmt.Heading3Bold);
            ApplySingleHeadingStyle("标题 4", "Heading 4", fmt.Heading4Font, fmt.Heading4Size, fmt.Heading4Bold);
        }

        private void ApplySingleHeadingStyle(string localName, string engName, string font, float size, bool bold)
        {
            try
            {
                Style style = null;
                try { style = _doc.Styles[localName]; }
                catch
                {
                    try { style = _doc.Styles[engName]; }
                    catch { return; }
                }
                if (style != null)
                {
                    style.Font.Name = font;
                    style.Font.Size = size;
                    style.Font.Bold = bold ? 1 : 0;
                }
            }
            catch { }
        }

        public void ApplyParagraphFormat(ParagraphFormat fmt)
        {
            foreach (Paragraph para in _doc.Paragraphs)
            {
                try
                {
                    Style style = (Style)para.get_Style();
                    if (IsHeadingStyle(style)) continue;
                    Range r = para.Range;
                    if ((int)r.get_Information(WdInformation.wdWithInTable) != 0) continue;

                    para.Format.Alignment = (WdParagraphAlignment)fmt.Alignment;
                    if (fmt.FirstLineIndent > 0)
                    {
                        para.Format.CharacterUnitFirstLineIndent = fmt.FirstLineIndent;
                    }
                    else
                    {
                        para.Format.CharacterUnitFirstLineIndent = 0;
                        para.Format.FirstLineIndent = 0;
                    }
                    para.Format.CharacterUnitLeftIndent = fmt.LeftIndent;
                    para.Format.CharacterUnitRightIndent = fmt.RightIndent;
                    para.Format.SpaceBefore = fmt.SpaceBefore;
                    para.Format.SpaceAfter = fmt.SpaceAfter;
                }
                catch { }
            }
        }

        private void ApplyBorder(Border border, BorderDef def)
        {
            if (border == null) return;
            try
            {
                if (def.Show)
                {
                    border.LineStyle = WdLineStyle.wdLineStyleSingle;
                    float w = def.Width;
                    int lwVal = 2;
                    if (w >= 1.5f) lwVal = 8;
                    else if (w >= 1f) lwVal = 6;
                    else if (w >= 0.75f) lwVal = 4;
                    border.LineWidth = (WdLineWidth)lwVal;
                }
                else
                {
                    border.LineStyle = WdLineStyle.wdLineStyleNone;
                }
            }
            catch { }
        }

        /// <summary>
        /// Safely set cell text, preserving the end-of-cell marker.
        /// </summary>
        private static void SetCellText(Range cellRange, string text)
        {
            Range r = cellRange.Duplicate;
            if (r.End > r.Start)
                r.End = r.End - 1; // exclude end-of-cell marker
            r.Text = text;
        }

        public void ApplyTableFormat(TableFormat fmt)
        {
            foreach (Table tbl in _doc.Tables)
            {
                try
                {
                    ApplyBorder(tbl.Borders[WdBorderType.wdBorderTop], fmt.TopBorder);
                    ApplyBorder(tbl.Borders[WdBorderType.wdBorderBottom], fmt.BottomBorder);
                    ApplyBorder(tbl.Borders[WdBorderType.wdBorderLeft], fmt.LeftBorder);
                    ApplyBorder(tbl.Borders[WdBorderType.wdBorderRight], fmt.RightBorder);
                    ApplyBorder(tbl.Borders[WdBorderType.wdBorderHorizontal], fmt.InsideHorizontalBorder);
                    ApplyBorder(tbl.Borders[WdBorderType.wdBorderVertical], fmt.InsideVerticalBorder);

                    if (fmt.AutoFit == 0)
                    {
                        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                    }
                    else if (fmt.AutoFit == 1)
                    {
                        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    }
                    else if (fmt.AutoFit == 2 && fmt.FixedWidth > 0)
                    {
                        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                        tbl.PreferredWidth = fmt.FixedWidth * 72f;
                    }

                    if (tbl.Rows.Count > 0)
                    {
                        Range headerRange = tbl.Rows[1].Range;
                        headerRange.Font.Name = fmt.HeaderFont;
                        headerRange.Font.Size = fmt.HeaderFontSize;
                        headerRange.Font.Bold = fmt.HeaderBold ? 1 : 0;
                    }

                    for (int i = 2; i <= tbl.Rows.Count; i++)
                    {
                        try
                        {
                            Range rowRange = tbl.Rows[i].Range;
                            rowRange.Font.Name = fmt.BodyFont;
                            rowRange.Font.Size = fmt.BodyFontSize;
                        }
                        catch { }
                    }

                    for (int r = 1; r <= tbl.Rows.Count; r++)
                    {
                        for (int c = 1; c <= tbl.Columns.Count; c++)
                        {
                            try
                            {
                                Range cellRange = tbl.Cell(r, c).Range;
                                string text = cellRange.Text;
                                // Remove end-of-cell marker (\r\a)
                                if (text.Length >= 2)
                                    text = text.Substring(0, text.Length - 2);
                                text = text.Trim();
                                double numVal;
                                if (double.TryParse(text, out numVal))
                                {
                                    cellRange.Font.Name = fmt.NumberFont;
                                    cellRange.Font.Size = fmt.NumberFontSize;
                                    if (fmt.UseThousandSeparator)
                                    {
                                        string formatted = numVal.ToString("N" + fmt.DecimalPlaces);
                                        SetCellText(cellRange, formatted);
                                    }
                                    else
                                    {
                                        string formatted = numVal.ToString("F" + fmt.DecimalPlaces);
                                        SetCellText(cellRange, formatted);
                                    }
                                    cellRange.ParagraphFormat.Alignment = (WdParagraphAlignment)fmt.NumberAlignment;
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
        }

        public void ApplyImageFormat(ImageFormat fmt)
        {
            foreach (InlineShape shape in _doc.InlineShapes)
            {
                try
                {
                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture ||
                        shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                    {
                        float maxW = fmt.MaxWidth * 72f;
                        float maxH = fmt.MaxHeight * 72f;
                        if (fmt.KeepAspectRatio)
                        {
                            float ratio = (float)shape.Width / (float)shape.Height;
                            if (shape.Width > maxW)
                            {
                                shape.Width = maxW;
                                shape.Height = maxW / ratio;
                            }
                            if (shape.Height > maxH)
                            {
                                shape.Height = maxH;
                                shape.Width = maxH * ratio;
                            }
                        }
                        else
                        {
                            if (shape.Width > maxW) shape.Width = maxW;
                            if (shape.Height > maxH) shape.Height = maxH;
                        }
                        shape.Range.ParagraphFormat.Alignment = (WdParagraphAlignment)fmt.Alignment;
                    }
                }
                catch { }
            }
        }

        public void ApplyNumberFormat(NumberFormat fmt)
        {
            foreach (Paragraph para in _doc.Paragraphs)
            {
                try
                {
                    Range r = para.Range;
                    // Only process standalone paragraphs (not in tables)
                    if ((int)r.get_Information(WdInformation.wdWithInTable) != 0) continue;
                    string text = r.Text.Trim();
                    double numVal;
                    if (double.TryParse(text, out numVal))
                    {
                        r.Font.Name = fmt.NumberFont;
                        r.Font.Size = fmt.NumberFontSize;
                        if (fmt.UseThousandSeparator)
                        {
                            r.Text = numVal.ToString("N" + fmt.DecimalPlaces);
                        }
                        else
                        {
                            r.Text = numVal.ToString("F" + fmt.DecimalPlaces);
                        }
                        para.Format.Alignment = (WdParagraphAlignment)fmt.Alignment;
                    }
                }
                catch { }
            }
        }

        public void ApplyHeadingLevel(int level)
        {
            try
            {
                Range sel = _doc.Application.Selection.Range;
                string styleName = "";
                switch (level)
                {
                    case 1: styleName = "标题 1"; break;
                    case 2: styleName = "标题 2"; break;
                    case 3: styleName = "标题 3"; break;
                    case 4: styleName = "标题 4"; break;
                    default: styleName = "正文"; break;
                }
                try
                {
                    sel.set_Style(_doc.Styles[styleName]);
                }
                catch
                {
                    try
                    {
                        string engName = "";
                        switch (level)
                        {
                            case 1: engName = "Heading 1"; break;
                            case 2: engName = "Heading 2"; break;
                            case 3: engName = "Heading 3"; break;
                            case 4: engName = "Heading 4"; break;
                            default: engName = "Normal"; break;
                        }
                        sel.set_Style(_doc.Styles[engName]);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
