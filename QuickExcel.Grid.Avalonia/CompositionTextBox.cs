using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;

namespace QuickExcel.Grid.Avalonia;

/// <summary>Observe the native TextBox preedit lifecycle without replacing its IME client.</summary>
public sealed class CompositionTextBox : TextBox
{
    protected override Type StyleKeyOverride => typeof(TextBox);
    private TextPresenter? _presenter;
    public bool IsComposing => !string.IsNullOrEmpty(_presenter?.PreeditText);
    public event Action? CompositionChanged;
    protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
    {
        if (_presenter is not null) _presenter.PropertyChanged -= PresenterChanged;
        base.OnApplyTemplate(e);
        _presenter = e.NameScope.Find<TextPresenter>("PART_TextPresenter");
        if (_presenter is not null) _presenter.PropertyChanged += PresenterChanged;
    }
    private void PresenterChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property == TextPresenter.PreeditTextProperty) CompositionChanged?.Invoke();
    }
}
