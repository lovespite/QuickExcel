using System.Collections.Immutable;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

[Flags]
public enum CellStyleParts { None = 0, Number = 1, Alignment = 2, Font = 4, Border = 8, Fill = 16, Protection = 32, All = 63 }

/// <summary>A reusable local format, independent of a workbook's style table and named-style IDs.</summary>
public sealed record QuickCellStyle(string Id, string Name, string Group, StyleDefinition Format, CellStyleParts Parts)
{
    public CellStylePatch Patch
    {
        get
        {
            var s = Format;
            bool Has(CellStyleParts part) => Parts.HasFlag(part);
            var font = Has(CellStyleParts.Font); var border = Has(CellStyleParts.Border);
            var alignment = Has(CellStyleParts.Alignment); var number = Has(CellStyleParts.Number); var protection = Has(CellStyleParts.Protection);
            return new(Bold: font ? s.Bold : null, Italic: font ? s.Italic : null,
                FontName: font ? s.FontName ?? "Microsoft YaHei UI" : null, FontSize: font ? s.FontSize ?? 11 : null,
                FontColor: font ? s.FontColor : null, ClearFontColor: font && s.FontColor is null,
                Underline: font ? s.Underline ?? "none" : null, Strike: font ? s.Strike : null,
                FillColor: Has(CellStyleParts.Fill) ? s.FillColor : null, ClearFillColor: Has(CellStyleParts.Fill) && s.FillColor is null,
                NumberFormatId: number ? s.NumberFormatId : null, NumberFormatCode: number ? s.NumberFormatCode : null,
                Horizontal: alignment ? s.Horizontal : null, Vertical: alignment ? s.Vertical : null,
                WrapText: alignment ? s.WrapText : null, TextRotation: alignment ? s.TextRotation : null,
                Indent: alignment ? s.Indent : null, ShrinkToFit: alignment ? s.ShrinkToFit : null,
                Locked: protection ? s.Locked : null, FormulaHidden: protection ? s.FormulaHidden : null,
                LeftBorder: border ? s.LeftBorder : null, RightBorder: border ? s.RightBorder : null,
                TopBorder: border ? s.TopBorder : null, BottomBorder: border ? s.BottomBorder : null,
                LeftBorderColor: border ? s.LeftBorderColor ?? "FF000000" : null, RightBorderColor: border ? s.RightBorderColor ?? "FF000000" : null,
                TopBorderColor: border ? s.TopBorderColor ?? "FF000000" : null, BottomBorderColor: border ? s.BottomBorderColor ?? "FF000000" : null,
                DiagonalBorder: border ? s.DiagonalBorder : null, DiagonalBorderColor: border ? s.DiagonalBorderColor ?? "FF000000" : null,
                DiagonalUp: border ? s.DiagonalUp : null, DiagonalDown: border ? s.DiagonalDown : null);
        }
    }

    // The format dialog edits a private one-cell document; previews never touch the open workbook or its history.
    public static StyleDefinition EditFormat(StyleDefinition format, CellStylePatch patch, CellBorderPatch? borders = null)
    {
        var id = SheetId.New();
        var sheet = new WorksheetSnapshot(id, "Style", new Dictionary<CellPosition, CellSnapshot> { [new(0, 0)] = new(CellContent.FromLiteral(CellValue.Blank)) });
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [format]));
        session.Execute(new SetCellFormatCommand([new(id, 0, 0)], patch, borders));
        return session.Snapshot.Styles[session.Snapshot.Sheets[0].GetStyleId(0, 0)] with { NamedStyleId = 0 };
    }
}

public static class QuickCellStyles
{
    private const CellStyleParts Appearance = CellStyleParts.Font | CellStyleParts.Fill | CellStyleParts.Border;
    public static ImmutableArray<QuickCellStyle> BuiltIn { get; } = Create();
    private static ImmutableArray<QuickCellStyle> Create()
    {
        var result = new List<QuickCellStyle>();
        void Add(string id, string name, string group, StyleDefinition style, CellStyleParts parts = Appearance)
            => result.Add(new(id, name, group, style, parts));
        Add("normal", "常规", "常用", new(), CellStyleParts.All);
        Add("bad", "差", "常用", new() { FontColor = "FF9C0006", FillColor = "FFFFC7CE" });
        Add("good", "好", "常用", new() { FontColor = "FF006100", FillColor = "FFC6EFCE" });
        Add("neutral", "适中", "常用", new() { FontColor = "FF9C6500", FillColor = "FFFFEB9C" });
        Add("calculation", "计算", "常用", new() { FontColor = "FFFA7D00", FillColor = "FFF2F2F2", LeftBorder = 1, RightBorder = 1, TopBorder = 1, BottomBorder = 1 });
        Add("check", "检查单元格", "常用", new() { FontColor = "FFFFFFFF", FillColor = "FFA5A5A5", Bold = true, LeftBorder = 6, RightBorder = 6, TopBorder = 6, BottomBorder = 6 });
        Add("explanation", "解释性文本", "常用", new() { FontColor = "FF7F7F7F", Italic = true });
        Add("warning", "警告文本", "常用", new() { FontColor = "FFFF0000" });
        Add("linked", "链接单元格", "常用", new() { FontColor = "FFFA7D00", BottomBorder = 6, BottomBorderColor = "FFFF8000" });
        Add("output", "输出", "常用", new() { Bold = true, FillColor = "FFF2F2F2", LeftBorder = 1, RightBorder = 1, TopBorder = 1, BottomBorder = 1 });
        Add("input", "输入", "常用", new() { FontColor = "FF3F3F76", FillColor = "FFFFCC99", LeftBorder = 1, RightBorder = 1, TopBorder = 1, BottomBorder = 1 });
        Add("title", "标题", "标题和总计", new() { FontSize = 18, Bold = true, FontColor = "FF1F4E78" });
        Add("heading1", "标题 1", "标题和总计", new() { FontSize = 15, Bold = true, FontColor = "FF1F4E78", BottomBorder = 5, BottomBorderColor = "FF4472C4" });
        Add("heading2", "标题 2", "标题和总计", new() { FontSize = 13, Bold = true, FontColor = "FF1F4E78", BottomBorder = 1, BottomBorderColor = "FF4472C4" });
        Add("heading3", "标题 3", "标题和总计", new() { Bold = true, FontColor = "FF1F4E78" });
        Add("heading4", "标题 4", "标题和总计", new() { Bold = true });
        Add("total", "总计", "标题和总计", new() { Bold = true, TopBorder = 1, BottomBorder = 6, TopBorderColor = "FF4472C4", BottomBorderColor = "FF4472C4" });
        string[] accents = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];
        for (var tone = 0; tone < 4; tone++)
            for (var index = 0; index < accents.Length; index++)
            {
                var rgb = accents[index]; var weight = new[] { 1d, .6, .4, .2 }[tone];
                var fill = "FF" + string.Concat(Enumerable.Range(0, 3).Select(c => ((int)Math.Round(255 + (Convert.ToInt32(rgb.Substring(c * 2, 2), 16) - 255) * weight)).ToString("X2")));
                var name = (tone == 0 ? "" : new[] { "", "60% - ", "40% - ", "20% - " }[tone]) + "强调文字颜色 " + (index + 1);
                Add($"accent{index + 1}-{tone}", name, "主题样式", new() { FillColor = fill, FontColor = tone < 2 ? "FFFFFFFF" : "FF" + rgb });
            }
        Add("comma", "千位分隔", "数字格式", new() { NumberFormatId = 4 }, CellStyleParts.Number);
        Add("comma0", "千位分隔 [0]", "数字格式", new() { NumberFormatId = 3 }, CellStyleParts.Number);
        Add("currency", "货币", "数字格式", new() { NumberFormatCode = "\"¥\"#,##0.00" }, CellStyleParts.Number);
        Add("currency0", "货币 [0]", "数字格式", new() { NumberFormatCode = "\"¥\"#,##0" }, CellStyleParts.Number);
        Add("percent", "百分比", "数字格式", new() { NumberFormatId = 9 }, CellStyleParts.Number);
        return result.ToImmutableArray();
    }

    public static string DisplayName(QuickCellStyle style)
    {
        if (!BuiltIn.Any(s => s.Id == style.Id)) return style.Name;
        if (style.Id.StartsWith("accent", StringComparison.Ordinal))
        {
            var suffix = style.Name.IndexOf('强');
            return style.Name[..suffix] + UiText.Format($"强调文字颜色 {style.Id[6]}");
        }
        return UiText.Get(style.Name);
    }
}
