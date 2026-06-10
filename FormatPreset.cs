using System;
using System.Collections.Generic;

namespace FormatHelper
{
    [Serializable]
    public class TextFormat
    {
        public string ChineseFont = "宋体";
        public string EnglishFont = "Times New Roman";
        public float FontSize = 12f;
        public bool Bold = false;
        public bool Italic = false;
        public float LineSpacing = 1.5f;
        public float SpaceBefore = 0f;
        public float SpaceAfter = 0f;
    }

    [Serializable]
    public class HeadingFormat
    {
        public string Heading1Font = "黑体";
        public float Heading1Size = 22f;
        public bool Heading1Bold = true;
        public string Heading2Font = "黑体";
        public float Heading2Size = 16f;
        public bool Heading2Bold = true;
        public string Heading3Font = "黑体";
        public float Heading3Size = 14f;
        public bool Heading3Bold = true;
        public string Heading4Font = "楷体";
        public float Heading4Size = 14f;
        public bool Heading4Bold = false;
    }

    [Serializable]
    public class ParagraphFormat
    {
        public int Alignment = 0;
        public float FirstLineIndent = 2f;
        public float LeftIndent = 0f;
        public float RightIndent = 0f;
        public float SpaceBefore = 0f;
        public float SpaceAfter = 0f;
    }

    [Serializable]
    public class BorderDef
    {
        public bool Show = true;
        public int Style = 0;
        public float Width = 0.5f;
    }

    [Serializable]
    public class TableFormat
    {
        public BorderDef TopBorder = new BorderDef();
        public BorderDef BottomBorder = new BorderDef();
        public BorderDef LeftBorder = new BorderDef();
        public BorderDef RightBorder = new BorderDef();
        public BorderDef InsideHorizontalBorder = new BorderDef();
        public BorderDef InsideVerticalBorder = new BorderDef();
        public int AutoFit = 1;
        public float FixedWidth = 0f;
        public string HeaderFont = "黑体";
        public float HeaderFontSize = 10.5f;
        public bool HeaderBold = true;
        public string BodyFont = "宋体";
        public float BodyFontSize = 10.5f;
        public string NumberFont = "Times New Roman";
        public float NumberFontSize = 10.5f;
        public int NumberAlignment = 0;
        public bool UseThousandSeparator = true;
        public int DecimalPlaces = 2;
    }

    [Serializable]
    public class ImageFormat
    {
        public float MaxWidth = 14f;
        public float MaxHeight = 20f;
        public bool KeepAspectRatio = true;
        public int Alignment = 1;
    }

    [Serializable]
    public class NumberFormat
    {
        public string NumberFont = "Times New Roman";
        public float NumberFontSize = 12f;
        public bool UseThousandSeparator = true;
        public int DecimalPlaces = 2;
        public int Alignment = 0;
    }

    [Serializable]
    public class FormatPreset
    {
        public string Name = "";
        public TextFormat Text = new TextFormat();
        public HeadingFormat Heading = new HeadingFormat();
        public ParagraphFormat Paragraph = new ParagraphFormat();
        public TableFormat Table = new TableFormat();
        public ImageFormat Image = new ImageFormat();
        public NumberFormat Number = new NumberFormat();
    }

    public static class PresetManager
    {
        public static List<FormatPreset> Presets = new List<FormatPreset>();

        static PresetManager()
        {
            Presets.Add(CreateStandardReport());
            Presets.Add(CreateAuditReport());
            Presets.Add(CreateCustomFormat());
        }

        public static FormatPreset GetPresetByName(string name)
        {
            foreach (FormatPreset p in Presets)
            {
                if (p.Name == name) return p;
            }
            return null;
        }

        private static FormatPreset CreateStandardReport()
        {
            FormatPreset p = new FormatPreset();
            p.Name = "标准报告";
            p.Text.ChineseFont = "宋体";
            p.Text.EnglishFont = "Times New Roman";
            p.Text.FontSize = 12f;
            p.Text.Bold = false;
            p.Text.LineSpacing = 1.5f;
            p.Text.SpaceBefore = 0f;
            p.Text.SpaceAfter = 0f;

            p.Heading.Heading1Font = "黑体";
            p.Heading.Heading1Size = 22f;
            p.Heading.Heading1Bold = true;
            p.Heading.Heading2Font = "黑体";
            p.Heading.Heading2Size = 16f;
            p.Heading.Heading2Bold = true;
            p.Heading.Heading3Font = "黑体";
            p.Heading.Heading3Size = 14f;
            p.Heading.Heading3Bold = true;
            p.Heading.Heading4Font = "楷体";
            p.Heading.Heading4Size = 14f;
            p.Heading.Heading4Bold = false;

            p.Paragraph.Alignment = 0;
            p.Paragraph.FirstLineIndent = 2f;
            p.Paragraph.LeftIndent = 0f;
            p.Paragraph.RightIndent = 0f;
            p.Paragraph.SpaceBefore = 0f;
            p.Paragraph.SpaceAfter = 0f;

            p.Table.HeaderFont = "黑体";
            p.Table.HeaderFontSize = 10.5f;
            p.Table.HeaderBold = true;
            p.Table.BodyFont = "宋体";
            p.Table.BodyFontSize = 10.5f;
            p.Table.NumberFont = "Times New Roman";
            p.Table.NumberFontSize = 10.5f;
            p.Table.UseThousandSeparator = true;
            p.Table.DecimalPlaces = 2;

            p.Image.MaxWidth = 14f;
            p.Image.MaxHeight = 20f;
            p.Image.KeepAspectRatio = true;
            p.Image.Alignment = 1;

            p.Number.NumberFont = "Times New Roman";
            p.Number.NumberFontSize = 12f;
            p.Number.UseThousandSeparator = true;
            p.Number.DecimalPlaces = 2;
            p.Number.Alignment = 0;

            return p;
        }

        private static FormatPreset CreateAuditReport()
        {
            FormatPreset p = new FormatPreset();
            p.Name = "审计报告";
            p.Text.ChineseFont = "仿宋";
            p.Text.EnglishFont = "Times New Roman";
            p.Text.FontSize = 16f;
            p.Text.Bold = false;
            p.Text.LineSpacing = 1.5f;
            p.Text.SpaceBefore = 0f;
            p.Text.SpaceAfter = 0f;

            p.Heading.Heading1Font = "方正小标宋简体";
            p.Heading.Heading1Size = 22f;
            p.Heading.Heading1Bold = true;
            p.Heading.Heading2Font = "黑体";
            p.Heading.Heading2Size = 16f;
            p.Heading.Heading2Bold = true;
            p.Heading.Heading3Font = "楷体";
            p.Heading.Heading3Size = 16f;
            p.Heading.Heading3Bold = false;
            p.Heading.Heading4Font = "仿宋";
            p.Heading.Heading4Size = 16f;
            p.Heading.Heading4Bold = false;

            p.Paragraph.Alignment = 0;
            p.Paragraph.FirstLineIndent = 2f;
            p.Paragraph.LeftIndent = 0f;
            p.Paragraph.RightIndent = 0f;
            p.Paragraph.SpaceBefore = 0f;
            p.Paragraph.SpaceAfter = 0f;

            p.Table.HeaderFont = "黑体";
            p.Table.HeaderFontSize = 10.5f;
            p.Table.HeaderBold = true;
            p.Table.BodyFont = "仿宋";
            p.Table.BodyFontSize = 10.5f;
            p.Table.NumberFont = "Times New Roman";
            p.Table.NumberFontSize = 10.5f;
            p.Table.UseThousandSeparator = true;
            p.Table.DecimalPlaces = 2;

            p.Image.MaxWidth = 14f;
            p.Image.MaxHeight = 20f;
            p.Image.KeepAspectRatio = true;
            p.Image.Alignment = 1;

            p.Number.NumberFont = "Times New Roman";
            p.Number.NumberFontSize = 12f;
            p.Number.UseThousandSeparator = true;
            p.Number.DecimalPlaces = 2;
            p.Number.Alignment = 0;

            return p;
        }

        private static FormatPreset CreateCustomFormat()
        {
            FormatPreset p = new FormatPreset();
            p.Name = "自定义格式";
            return p;
        }
    }
}
