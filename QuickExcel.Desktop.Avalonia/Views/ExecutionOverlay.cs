using System.Diagnostics;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Keeps short operations visually quiet; input and command guards remain active.</summary>
internal sealed class ExecutionOverlay : global::Avalonia.Controls.Grid
{
    private readonly Window _owner;
    private readonly Control _content;
    private readonly Func<bool> _isComposing;
    private readonly Action<bool>? _visibilityChanged;
    private readonly Border _overlay;
    private readonly TextBlock _title = new() { FontSize = 16, FontWeight = FontWeight.SemiBold, Foreground = Brushes.Black, HorizontalAlignment = HorizontalAlignment.Center };
    internal TextBlock Elapsed { get; } = new() { FontFamily = new("Consolas"), Foreground = Brushes.Black, HorizontalAlignment = HorizontalAlignment.Center };
    internal Button Cancel { get; } = new() { Content = UiText.Get("取消  Esc"), HorizontalAlignment = HorizontalAlignment.Center, MinWidth = 112 };
    private readonly RotateTransform _rotation = new();
    private readonly DispatcherTimer _timer = new();
    private readonly Stopwatch _elapsed = new();
    private CancellationTokenSource? _cancellation;
    private IInputElement? _previousFocus;
    private bool _contentHitTestVisible;
    internal bool IsRunning => _cancellation is not null;
    internal bool IsOverlayVisible => _overlay.IsVisible;

    internal ExecutionOverlay(Window owner, Control content, Func<bool> isComposing, Action<bool>? visibilityChanged = null)
    {
        _owner = owner; _content = content; _isComposing = isComposing; _visibilityChanged = visibilityChanged;
        var spinner = new global::Avalonia.Controls.Shapes.Path
        {
            Width = 36, Height = 36, HorizontalAlignment = HorizontalAlignment.Center,
            Data = Geometry.Parse("M 18,3 A 15,15 0 1 1 3,18"), Stroke = Brush.Parse("#1967D2"), StrokeThickness = 3,
            RenderTransform = _rotation, RenderTransformOrigin = RelativePoint.Center
        };
        AutomationProperties.SetName(spinner, UiText.Get("正在加载"));
        AutomationProperties.SetName(Cancel, UiText.Get("取消执行（Esc）"));
        var panel = new StackPanel { Spacing = 14 };
        panel.Children.Add(spinner); panel.Children.Add(_title); panel.Children.Add(Elapsed); panel.Children.Add(Cancel);
        _overlay = new Border
        {
            IsVisible = false, Background = Brush.Parse("#66000000"),
            Child = new Border { Background = Brushes.White, CornerRadius = new(10), Padding = new(36, 28), MinWidth = 250,
                HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Child = panel }
        };
        Children.Add(content); Children.Add(_overlay);
        Cancel.Click += (_, _) => CancelExecution();
        _timer.Tick += (_, _) => Update();
        owner.AddHandler(KeyDownEvent, OnKeyDown, RoutingStrategies.Tunnel);
        owner.AddHandler(TextInputEvent, (_, e) => { if (IsRunning) e.Handled = true; }, RoutingStrategies.Tunnel);
        owner.Activated += (_, _) => { if (_overlay.IsVisible) Cancel.Focus(); };
        owner.Closed += (_, _) => { _cancellation?.Cancel(); End(); };
    }

    internal IDisposable Begin(string title, CancellationTokenSource cancellation)
    {
        if (IsRunning) throw new InvalidOperationException(UiText.Get("已有计算或脚本正在执行。"));
        _cancellation = cancellation; _title.Text = title; Cancel.Content = UiText.Get("取消  Esc"); Cancel.IsEnabled = true;
        // Do not toggle IsEnabled: Fluent's disabled styles flash even when work finishes in one frame.
        _contentHitTestVisible = _content.IsHitTestVisible; _content.IsHitTestVisible = false;
        _elapsed.Restart(); _timer.Interval = TimeSpan.FromMilliseconds(100); _timer.Start();
        return new Scope(this, cancellation);
    }

    private void Update()
    {
        if (!IsRunning || _elapsed.ElapsedMilliseconds < 100) return;
        Elapsed.Text = $"{_elapsed.ElapsedMilliseconds:N0} ms";
        _rotation.Angle = _elapsed.Elapsed.TotalMilliseconds * .36 % 360;
        if (_overlay.IsVisible) return;
        _previousFocus = _owner.FocusManager?.GetFocusedElement();
        _visibilityChanged?.Invoke(true); _overlay.IsVisible = true;
        _timer.Interval = TimeSpan.FromMilliseconds(33);
        if (_owner.IsActive) Cancel.Focus();
    }

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        if (!IsRunning || e.Handled) return;
        if (e.Key == Key.Escape && !_isComposing()) { CancelExecution(); e.Handled = true; return; }
        // The cancel button owns its normal keyboard activation; all underlying editors stay unchanged.
        if (!_overlay.IsVisible || !Cancel.IsKeyboardFocusWithin || e.Key is not (Key.Enter or Key.Space)) e.Handled = true;
    }

    private void CancelExecution()
    {
        _cancellation?.Cancel(); Cancel.Content = UiText.Get("正在取消…");
    }

    private void End()
    {
        if (!IsRunning) return;
        _timer.Stop(); _elapsed.Stop(); _cancellation = null;
        var restoreFocus = _overlay.IsKeyboardFocusWithin;
        _overlay.IsVisible = false; _content.IsHitTestVisible = _contentHitTestVisible;
        _visibilityChanged?.Invoke(false);
        if (restoreFocus && _owner.IsActive) _previousFocus?.Focus();
        _previousFocus = null;
    }

    private sealed class Scope(ExecutionOverlay owner, CancellationTokenSource cancellation) : IDisposable
    {
        public void Dispose() { if (ReferenceEquals(owner._cancellation, cancellation)) owner.End(); }
    }
}
