using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace QuickExcel.Desktop;

/// <summary>Let the document resize path own header grips before WPF mutates their geometry.</summary>
internal static class GridResizeInput
{
    private static readonly DependencyProperty EnabledProperty = DependencyProperty.RegisterAttached(
        "Enabled", typeof(bool), typeof(GridResizeInput), new PropertyMetadata(false));

    static GridResizeInput()
    {
        // Class handlers precede the native header's handlers attached to each Thumb.
        // The grid still receives the events through its handledEventsToo subscriptions.
        EventManager.RegisterClassHandler(typeof(Thumb), Thumb.DragStartedEvent, new DragStartedEventHandler(Handle));
        EventManager.RegisterClassHandler(typeof(Thumb), Thumb.DragDeltaEvent, new DragDeltaEventHandler(Handle));
        EventManager.RegisterClassHandler(typeof(Thumb), Thumb.DragCompletedEvent, new DragCompletedEventHandler(Handle));
        EventManager.RegisterClassHandler(typeof(Thumb), Control.MouseDoubleClickEvent, new MouseButtonEventHandler(Handle));
    }

    public static void Enable(DataGrid grid) => grid.SetValue(EnabledProperty, true);

    private static void Handle(object sender, RoutedEventArgs e)
    {
        if (sender is not Thumb { TemplatedParent: DataGridColumnHeader or DataGridRowHeader } grip ||
            grip.Name is not ("PART_LeftHeaderGripper" or "PART_RightHeaderGripper" or "PART_TopHeaderGripper" or "PART_BottomHeaderGripper")) return;
        for (DependencyObject? parent = grip.TemplatedParent; parent is not null; parent = VisualTreeHelper.GetParent(parent))
            if (parent is DataGrid grid)
            {
                if ((bool)grid.GetValue(EnabledProperty)) e.Handled = true;
                return;
            }
    }
}
