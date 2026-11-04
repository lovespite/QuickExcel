using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private static readonly Dictionary<string, global::Avalonia.Media.Imaging.Bitmap> ConditionalIconSamples = [];
    private static Control IconSetSample(string name)
    {
        if (!ConditionalIconSamples.TryGetValue(name, out var image))
        {
            var count = name[0] - '0'; using var bitmap = new SKBitmap(count * 22, 22); using var canvas = new SKCanvas(bitmap); canvas.Clear(SKColors.Transparent);
            for (var i = 0; i < count; i++) SkiaRenderer.DrawConditionalIcon(canvas, new(name, count - 1 - i), SKRect.Create(i * 22 + 2, 2, 18, 18));
            using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100); using var stream = encoded.AsStream();
            ConditionalIconSamples[name] = image = new global::Avalonia.Media.Imaging.Bitmap(stream);
        }
        return new Image { Source = image, Width = 110, Height = 22, Stretch = Stretch.None, HorizontalAlignment = HorizontalAlignment.Left };
    }
    private static Control ColorScaleSample(IEnumerable<string> colors)
    {
        var values = colors.ToArray();
        var brush = new LinearGradientBrush { StartPoint = new(0, .5, RelativeUnit.Relative), EndPoint = new(1, .5, RelativeUnit.Relative) };
        for (var i = 0; i < values.Length; i++) brush.GradientStops.Add(new(Color.Parse("#" + values[i]), (double)i / (values.Length - 1)));
        return new Border { Width = 110, Height = 20, Background = brush, BorderBrush = Brushes.Gray, BorderThickness = new(.5) };
    }
    private MenuItem VisualMenuItem(string name, Control sample, Func<CellRectangle, ConditionalFormatRule> create)
    {
        var item = CommandMenu(name, () => ApplyVisualPresetAsync(create));
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 }; row.Children.Add(sample);
        row.Children.Add(new TextBlock { Text = name, VerticalAlignment = VerticalAlignment.Center });
        item.Header = row; global::Avalonia.Automation.AutomationProperties.SetName(item, name); item.MinHeight = 34; return item;
    }
    private MenuItem DataBarMenu()
    {
        MenuItem Fill(bool gradient)
        {
            var group = new MenuItem { Header = gradient ? "渐变填充" : "实心填充" };
            foreach (var (name, color) in ConditionalVisualPresets.BarColors)
                group.Items.Add(VisualMenuItem(name, ColorScaleSample(gradient ? [color, "FFFFFFFF"] : [color, color]), range => ConditionalVisualPresets.DataBar(range, color, gradient)));
            return group;
        }
        return new() { Header = "数据条(D)", Icon = new ConditionalFormatIcon(2), MinHeight = 44, Items = { Fill(true), Fill(false), new Separator(),
            CommandMenu("其他数据条规则…", () => NewVisualRuleAsync(r => ConditionalVisualPresets.DataBar(r, ConditionalVisualPresets.BarColors[0].Color))) } };
    }
    private MenuItem ColorScaleMenu()
    {
        var group = new MenuItem { Header = "色阶(S)", Icon = new ConditionalFormatIcon(3), MinHeight = 44 };
        foreach (var (name, colors) in ConditionalVisualPresets.Scales)
            group.Items.Add(VisualMenuItem(name, ColorScaleSample(colors.Reverse()), range => ConditionalVisualPresets.ColorScale(range, colors)));
        group.Items.Add(new Separator());
        group.Items.Add(CommandMenu("其他双色规则…", () => NewVisualRuleAsync(r => ConditionalVisualPresets.ColorScale(r, ConditionalVisualPresets.Scales[6].Colors))));
        group.Items.Add(CommandMenu("其他三色规则…", () => NewVisualRuleAsync(r => ConditionalVisualPresets.ColorScale(r, ConditionalVisualPresets.Scales[0].Colors))));
        return group;
    }
    private MenuItem IconSetMenu()
    {
        var group = new MenuItem { Header = "图标集(I)", Icon = new ConditionalFormatIcon(4), MinHeight = 44 };
        foreach (var indices in new[] { new[] { 0, 1, 8, 9, 13, 14 }, new[] { 2, 3, 4, 5, 10, 12 }, new[] { 6, 7 }, new[] { 11, 15, 16 } })
        {
            if (group.Items.Count > 0) group.Items.Add(new Separator());
            foreach (var i in indices)
            {
                var name = ConditionalIconSet.Names[i];
                group.Items.Add(VisualMenuItem(ConditionalVisualPresets.IconLabels[i], IconSetSample(name), range => ConditionalVisualPresets.IconSet(range, name)));
            }
        }
        group.Items.Add(new Separator()); group.Items.Add(CommandMenu("其他图标集规则…", () => NewVisualRuleAsync(r => ConditionalVisualPresets.IconSet(r, "3Arrows")))); return group;
    }
    private async Task ApplyVisualPresetAsync(Func<CellRectangle, ConditionalFormatRule> create)
    {
        if (!await FlushEditsAsync()) return; RequireConditionalFormatEditing();
        var snapshot = Document.Snapshot; await AddConditionalFormatAsync(snapshot, create(Rectangle(SelectedRange())));
    }
    private async Task NewVisualRuleAsync(Func<CellRectangle, ConditionalFormatRule> create)
    {
        if (!await FlushEditsAsync()) return; RequireConditionalFormatEditing();
        var snapshot = Document.Snapshot; var old = create(Rectangle(SelectedRange())); var dialog = VisualRuleFields(old, out var read);
        dialog.Validate = () => read().ValidateBinding(snapshot, snapshot.ActiveSheetId);
        if (await dialog.ShowAsync(this)) await AddConditionalFormatAsync(snapshot, read());
    }
    private static FieldsDialog VisualRuleFields(ConditionalFormatRule old, out Func<ConditionalFormatRule> read)
    {
        var dialog = new FieldsDialog(ConditionalKinds[(int)old.Kind] + "规则").Text("ranges", "应用区域（多个区域用空格分隔）", string.Join(' ', old.Ranges));
        var thresholds = old.Thresholds.ToArray();
        if (old.IconSet is { } icons)
            dialog.Choice("icons", "图标样式", ConditionalVisualPresets.IconLabels.ToArray(), ConditionalIconSet.Names.IndexOf(icons.Name)).Check("reverse", "反转图标顺序", icons.Reverse).Check("show", "同时显示单元格值", icons.ShowValue);
        if (old.DataBar is { } bar)
        {
            dialog.Choice("fill", "填充方式", ["渐变填充", "实心填充"], bar.Gradient ? 0 : 1).Text("color", "正值颜色（ARGB）", bar.Color).Text("negative", "负值颜色（ARGB）", bar.NegativeColor)
                .Check("show", "同时显示单元格值", bar.ShowValue).Choice("axis", "坐标轴", ["自动", "居中", "无"], (int)bar.Axis)
                .Choice("direction", "方向", ["随工作表", "从左到右", "从右到左"], (int)bar.Direction).Text("axisColor", "坐标轴颜色（ARGB）", bar.AxisColor)
                .Text("border", "边框颜色（ARGB，留空无边框）", bar.BorderColor ?? "").Text("negativeBorder", "负值边框颜色（留空沿用正值）", bar.NegativeBorderColor ?? "")
                .Text("minLength", "最短长度（0–100%）", bar.MinLength.ToString(CultureInfo.InvariantCulture)).Text("maxLength", "最长长度（0–100%）", bar.MaxLength.ToString(CultureInfo.InvariantCulture));
        }
        dialog.Note("阈值公式使用绝对单元格引用，例如 =$D$1。阈值按从小到大的顺序填写。");
        var kinds = old.IconSet is not null ? new[] { ConditionalThresholdKind.Number, ConditionalThresholdKind.Percent, ConditionalThresholdKind.Percentile, ConditionalThresholdKind.Formula } :
            old.DataBar is not null ? Enum.GetValues<ConditionalThresholdKind>() : Enum.GetValues<ConditionalThresholdKind>().Take(6).ToArray();
        string[] labels = ["最低值", "最高值", "数字", "百分比", "百分位", "公式", "自动最小值", "自动最大值"];
        var fieldCount = old.IconSet is null ? thresholds.Length : 5;
        for (var i = 0; i < fieldCount; i++)
        {
            var last = thresholds[^1];
            var t = i < thresholds.Length ? thresholds[i] : last.Kind is ConditionalThresholdKind.Percent or ConditionalThresholdKind.Percentile ?
                last with { Value = last.Value + (100 - last.Value) * (i - thresholds.Length + 1) / (6 - thresholds.Length) } : last;
            var label = old.IconSet is null ? i == 0 ? "最小值" : i == thresholds.Length - 1 ? "最大值" : "中间值" : $"阈值 {i + 1}";
            dialog.Choice("type" + i, label + "类型", kinds.Select(k => labels[(int)k]).ToArray(), Array.IndexOf(kinds, t.Kind))
                .Text("value" + i, label + "数值或公式", t.Formula ?? t.Value.ToString(CultureInfo.CurrentCulture));
            if (old.IconSet is not null) dialog.Check("gte" + i, label + "包含边界（≥）", t.Inclusive);
            if (old.ColorScale is { } scale) dialog.Text("color" + i, label + "颜色（ARGB）", scale.Colors[i]);
        }
        if (old.IconSet is not null) dialog.Note("图标样式有 3、4 或 5 档，仅使用对应数量的阈值；第一阈值为最低档起点。");
        read = () =>
        {
            string Color(string key) => dialog.String(key).Trim().TrimStart('#').ToUpperInvariant() is var text && text.Length == 6 ? "FF" + text : text;
            ConditionalThreshold Threshold(int i)
            {
                var kind = kinds[dialog.Index("type" + i)]; var formula = kind == ConditionalThresholdKind.Formula ? dialog.String("value" + i) : null;
                var value = kind is ConditionalThresholdKind.Number or ConditionalThresholdKind.Percent or ConditionalThresholdKind.Percentile ? dialog.Number("value" + i, -double.MaxValue, double.MaxValue) : 0;
                return new(kind, value, formula, old.IconSet is null || dialog.Boolean("gte" + i));
            }
            var ranges = dialog.String("ranges").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse).ToArray();
            ConditionalColorScale? scale = old.ColorScale is null ? null : new(Enumerable.Range(0, thresholds.Length).Select(Threshold), Enumerable.Range(0, thresholds.Length).Select(i => Color("color" + i)));
            var data = old.DataBar is null ? null : new ConditionalDataBar(Threshold(0), Threshold(1), Color("color"), dialog.Index("fill") == 0, dialog.Boolean("show"),
                dialog.Integer("minLength", 0, 100), dialog.Integer("maxLength", 0, 100), (DataBarAxis)dialog.Index("axis"), (DataBarDirection)dialog.Index("direction"), Color("negative"), Color("axisColor"),
                string.IsNullOrWhiteSpace(dialog.String("border")) ? null : Color("border"), string.IsNullOrWhiteSpace(dialog.String("negativeBorder")) ? null : Color("negativeBorder"));
            var iconName = old.IconSet is null ? null : ConditionalIconSet.Names[dialog.Index("icons")];
            var icon = iconName is null ? null : new ConditionalIconSet(iconName, Enumerable.Range(0, iconName[0] - '0').Select(Threshold), dialog.Boolean("show"), dialog.Boolean("reverse"));
            return new ConditionalFormatRule(ranges, old.Kind, old.Priority, new(), stopIfTrue: old.StopIfTrue, colorScale: scale, dataBar: data, iconSet: icon) { ScriptId = old.ScriptId };
        };
        return dialog;
    }
}
