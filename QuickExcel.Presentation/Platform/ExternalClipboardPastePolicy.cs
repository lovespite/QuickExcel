using QuickExcel.Core;

namespace QuickExcel.Presentation;

public sealed record ExternalClipboardPastePlan(CellPasteMode Mode, string? Notice);

/// <summary>Only ordinary pastes of identified filter regions may automatically use source values.</summary>
public static class ExternalClipboardPastePolicy
{
    public static ExternalClipboardPastePlan Prepare(ExternalClipboardRegion source, CellPasteMode requested)
    {
        var fallback = requested == CellPasteMode.All && source.HasFilterMetadata && !source.Limitations.IsEmpty;
        var mode = fallback ? CellPasteMode.Values : requested;
        // Missing results and corrupt input are never converted to text or empty cells.
        source.RequireMode(mode);
        var notice = !source.HasFilterMetadata ? null : fallback
            ? "已自动改为只贴值，保留目标格式；来源筛选按钮和规则未复制。" + string.Join("；", source.Limitations)
            : "来源筛选按钮和条件未复制，已按普通单元格粘贴。";
        return new(mode, notice);
    }
}
