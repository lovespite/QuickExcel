using System.Windows;
using FluentIcons.Common;

namespace QuickExcel.Desktop;

/// <summary>Icon metadata shared by ribbon buttons and toggle buttons.</summary>
public static class Ribbon
{
    public static readonly DependencyProperty IconProperty = DependencyProperty.RegisterAttached(
        "Icon", typeof(Icon), typeof(Ribbon), new PropertyMetadata(default(Icon)));

    public static Icon GetIcon(DependencyObject element) => (Icon)element.GetValue(IconProperty);
    public static void SetIcon(DependencyObject element, Icon value) => element.SetValue(IconProperty, value);
}
