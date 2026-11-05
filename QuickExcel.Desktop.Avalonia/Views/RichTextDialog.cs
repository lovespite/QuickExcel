using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class RichTextDialog : Window
{
    internal CompositionTextBox Editor { get; } = new() { AcceptsReturn = true, NewLine = "\n", TextWrapping = TextWrapping.Wrap, MinHeight = 150, MaxHeight = 240, IsUndoEnabled = false };
    private readonly Image _preview = new() { Width = 650, Height = 150, Stretch = Stretch.Uniform };
    private readonly TextBlock _error = new() { TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Firebrick };
    private readonly StyleDefinition _style;
    private readonly Stack<RichCellText> _undo = new(), _redo = new();
    private bool _setting;
    private Bitmap? _bitmap;
    internal RichCellText Draft { get; private set; }

    public RichTextDialog(RichCellText source, StyleDefinition style)
    {
        Draft = source; _style = style; Editor.Text = source.Text;
        Title = UiText.Get("富文本编辑"); Width = 730; SizeToContent = SizeToContent.Height; MaxHeight = 800; WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var panel = new StackPanel { Spacing = 12, Margin = new(20) }; Content = panel;
        panel.Children.Add(new TextBlock { Text = UiText.Get("在文字框选中文字后设置格式。下方预览使用网格/打印的文本引擎；超出预览范围的内容仍完整保存。确定后整次编辑可一次撤销。"), TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(Editor);
        var tools = new WrapPanel { Orientation = Orientation.Horizontal }; panel.Children.Add(tools);
        void Button(string title, Action action) { var button = new Button { Content = title, Margin = new(0, 0, 8, 8) }; button.Click += (_, _) => Guard(action); tools.Children.Add(button); }
        Button(UiText.Get("粗体"), () => Apply(new(Bold: !(CurrentFont()?.Bold ?? _style.Bold))));
        Button(UiText.Get("斜体"), () => Apply(new(Italic: !(CurrentFont()?.Italic ?? _style.Italic))));
        var name = new TextBox { PlaceholderText = UiText.Get("字体名称"), Width = 150, Margin = new(0, 0, 8, 8) }; tools.Children.Add(name);
        Button(UiText.Get("应用字体"), () => Apply(new(Name: name.Text?.Trim() ?? "")));
        var size = new NumericUpDown { Minimum = 1, Maximum = 409, Value = (decimal)(style.FontSize ?? 11), Width = 90, Margin = new(0, 0, 8, 8) }; tools.Children.Add(size);
        Button(UiText.Get("应用字号"), () => Apply(new(Size: (double)(size.Value ?? 11))));
        var color = new TextBox { PlaceholderText = "#RRGGBB", Width = 100, Margin = new(0, 0, 8, 8) }; tools.Children.Add(color);
        Button(UiText.Get("应用颜色"), () => { var rgb = (color.Text ?? "").Trim().TrimStart('#'); if (rgb.Length == 6) rgb = "FF" + rgb; Apply(new(Color: rgb.ToUpperInvariant())); });
        Button(UiText.Get("撤销"), () => History(false)); Button(UiText.Get("重做"), () => History(true));
        panel.Children.Add(_preview); panel.Children.Add(_error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8 };
        var ok = new Button { Content = UiText.Get("确定") }; var cancel = new Button { Content = UiText.Get("取消"), IsCancel = true };
        ok.Click += (_, _) => Guard(() => { if (!Editor.IsComposing) { SyncText(); Close(true); } }); cancel.Click += (_, _) => { if (!Editor.IsComposing) Close(false); };
        buttons.Children.Add(ok); buttons.Children.Add(cancel); panel.Children.Add(buttons);
        Editor.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty && !_setting && !Editor.IsComposing) Guard(SyncText); };
        Editor.CompositionChanged += () => { if (!Editor.IsComposing) Guard(SyncText); };
        AddHandler(KeyDownEvent, (_, e) =>
        {
            if (Editor.IsComposing) return;
            if ((e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) != 0 && e.Key is Key.Z or Key.Y)
            { e.Handled = true; Guard(() => History(e.Key == Key.Y || e.KeyModifiers.HasFlag(KeyModifiers.Shift))); }
        }, RoutingStrategies.Tunnel);
        Closing += (_, e) => { if (Editor.IsComposing) e.Cancel = true; };
        Opened += (_, _) => { RefreshPreview(); Editor.Focus(); }; Closed += (_, _) => _bitmap?.Dispose();
    }
    private TextRunFont? CurrentFont() => RichTextEditing.FontAt(Draft, Math.Min(Editor.SelectionStart, Editor.SelectionEnd));
    private void Guard(Action action) { try { action(); _error.Text = ""; } catch (Exception ex) { _error.Text = ex.Message; } }
    internal void Apply(TextRunFont patch)
    {
        if (Editor.IsComposing) return; SyncText();
        var start = Math.Min(Editor.SelectionStart, Editor.SelectionEnd); var length = Math.Abs(Editor.SelectionEnd - Editor.SelectionStart);
        Change(RichTextEditing.Format(Draft, start, length, patch, _style));
    }
    private void Change(RichCellText next)
    {
        if (Draft.Equals(next)) return; _undo.Push(Draft); _redo.Clear(); Draft = next;
        if (_undo.Count > 100) { var newest = _undo.Take(100).Reverse().ToArray(); _undo.Clear(); foreach (var state in newest) _undo.Push(state); }
        RefreshPreview();
    }
    private void SyncText()
    {
        var text = Editor.Text ?? ""; var before = Draft.Text; if (text == before) return;
        var start = 0; while (start < text.Length && start < before.Length && text[start] == before[start]) start++;
        if (start > 0 && start < before.Length && char.IsHighSurrogate(before[start - 1])) start--;
        var end = 0; while (end < before.Length - start && end < text.Length - start && before[^(end + 1)] == text[^(end + 1)]) end++;
        if (end > 0 && end < before.Length && char.IsLowSurrogate(before[^end])) end--;
        Change(RichTextEditing.Replace(Draft, start, before.Length - start - end, text.Substring(start, text.Length - start - end)));
    }
    private void History(bool redo)
    {
        if (Editor.IsComposing) return; var source = redo ? _redo : _undo; var target = redo ? _undo : _redo;
        if (!source.TryPop(out var state)) return; target.Push(Draft); Draft = state;
        _setting = true; Editor.Text = Draft.Text; _setting = false; RefreshPreview();
    }
    private void RefreshPreview()
    {
        using var renderer = new SkiaRenderer(); using var surface = SKSurface.Create(new SKImageInfo(650, 150));
        surface.Canvas.Clear(SKColors.White);
        renderer.DrawText(surface.Canvas, new(Draft.Text, _style with { WrapText = true, Vertical = 0, Horizontal = 1, TextRotation = 0 }, RichText: Draft), SKRect.Create(6, 6, 638, 138));
        using var image = surface.Snapshot(); using var png = image.Encode(SKEncodedImageFormat.Png, 100); using var stream = png.AsStream();
        var previous = _bitmap; _bitmap = new Bitmap(stream); _preview.Source = _bitmap; previous?.Dispose();
    }
}
