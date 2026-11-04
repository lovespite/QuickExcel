using QuickExcel.Core;

namespace QuickExcel.Presentation;

public sealed record ExternalClipboardPastePlan(CellPasteMode Mode, string? Notice);

/// <summary>Paste the modeled portion of a source and explain any omitted content.</summary>
public static class ExternalClipboardPastePolicy
{
    public static ExternalClipboardPastePlan Prepare(ExternalClipboardRegion source, CellPasteMode requested)
    {
        if (!Enum.IsDefined(requested)) throw new ArgumentException("未知粘贴模式。");
        var mode = source.ValuesAreDisplayText && requested == CellPasteMode.Formulas ? CellPasteMode.Values : requested;
        var notices = new List<string>();
        if (source.HasFilterMetadata) notices.Add("来源筛选按钮和条件未复制，已按普通单元格粘贴。");
        if (!source.Limitations.IsEmpty) notices.Add("已粘贴支持的部分；未保留的内容：" + string.Join("；", source.Limitations) + "。");
        if (source.ValuesAreDisplayText && mode != CellPasteMode.Formats) notices.Add("来源仅提供显示文本，已按文本值粘贴。");
        if (!source.HasFormats && mode is CellPasteMode.All or CellPasteMode.Formats) notices.Add("缺少来源格式的单元格保留目标格式。");
        if (mode != CellPasteMode.Formats && source.Cells.Any(c => c.MissingResult)) notices.Add("缺少可用公式结果的单元格已跳过，保留目标内容及合并结构。");
        if (mode == CellPasteMode.Formulas && source.Cells.Any(c => c.FormulaRejection is not null)) notices.Add("未支持的公式已改贴结果值。");
        return new(mode, notices.Count == 0 ? null : string.Join(" ", notices));
    }
}
