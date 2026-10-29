using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task MeasureResizeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        var cells = Enumerable.Range(0, 10_000).ToDictionary(i => new CellPosition(i / 10, i % 10),
            i => new CellSnapshot(i % 10 == 9 ? CellContent.FromFormula($"=A{i / 10 + 1}+1") : CellContent.FromLiteral(CellValue.FromNumber(i))));
        InstallDocument(new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "Resize", cells)]), _services), null, "尺寸性能测量.xlsq");
        await RecalculateAsync();
        GridSurface.Select(new(2, 2)); Freeze_Click(new MenuItem(), new());
        await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
        var samples = new List<object>();
        var recalculations = 0; var disabled = 0; var frozenResets = 0;
        void Changed(object? sender, DocumentChange e) { if (e.CalculationOnly) recalculations++; }
        void EnabledChanged(object sender, DependencyPropertyChangedEventArgs e) { if (Equals(e.NewValue, false)) disabled++; }
        void ColumnsChanged(object? sender, NotifyCollectionChangedEventArgs e) { if (e.Action == NotifyCollectionChangedAction.Reset) frozenResets++; }
        _document.Session.Changed += Changed;
        Workspace.IsEnabledChanged += EnabledChanged;
        ((INotifyCollectionChanged)FrozenGrid.Columns).CollectionChanged += ColumnsChanged;
        var suppress = _suppressAutoCalculation;
        _suppressAutoCalculation = false;
        try
        {
            foreach (var axis in new[] { "column", "row" })
            for (var sample = 0; sample < 15; sample++)
            {
                Thumb grip;
                if (axis == "column")
                {
                    var presenter = VisualChild<DataGridColumnHeadersPresenter>(SheetGrid)!;
                    var header = (DataGridColumnHeader)presenter.ItemContainerGenerator.ContainerFromIndex(0);
                    grip = (Thumb)header.Template.FindName("PART_RightHeaderGripper", header);
                }
                else
                {
                    var row = (DataGridRow)SheetGrid.ItemContainerGenerator.ContainerFromItem(_rows[0]);
                    var header = VisualChild<DataGridRowHeader>(row)!;
                    grip = (Thumb)header.Template.FindName("PART_BottomHeaderGripper", header);
                }
                recalculations = disabled = frozenResets = 0;
                var preview = new List<double>();
                grip.RaiseEvent(new DragStartedEventArgs(0, 0) { RoutedEvent = Thumb.DragStartedEvent });
                for (var frame = 0; frame < 4; frame++)
                {
                    var timer = Stopwatch.StartNew();
                    for (var move = 0; move < 8; move++)
                    {
                        var delta = sample % 2 == 0 ? 0.5 : -0.5;
                        // Thumb offsets overlap until layout moves the grip to the pointer.
                        grip.RaiseEvent(new DragDeltaEventArgs(axis == "column" ? delta * (move + 1) : 0, axis == "row" ? delta * (move + 1) : 0) { RoutedEvent = Thumb.DragDeltaEvent });
                    }
                    await Dispatcher.InvokeAsync(() => SheetGrid.UpdateLayout(), DispatcherPriority.ApplicationIdle);
                    preview.Add(timer.Elapsed.TotalMilliseconds);
                }
                var complete = Stopwatch.StartNew();
                grip.RaiseEvent(new DragCompletedEventArgs(0, 0, false) { RoutedEvent = Thumb.DragCompletedEvent });
                await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
                while (_busy) await Task.Delay(5);
                await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
                if (sample >= 5) samples.Add(new { axis, previewMs = preview, completeMs = complete.Elapsed.TotalMilliseconds, recalculations, workspaceDisableCount = disabled, frozenColumnResets = frozenResets });
            }
        }
        finally
        {
            _suppressAutoCalculation = suppress;
            _document.Session.Changed -= Changed;
            Workspace.IsEnabledChanged -= EnabledChanged;
            ((INotifyCollectionChanged)FrozenGrid.Columns).CollectionChanged -= ColumnsChanged;
        }
        File.WriteAllText(Path.Combine(folder, "resize-benchmark.json"), JsonSerializer.Serialize(new
        {
            method = "10k cells including 1k formulas, 2 frozen rows/columns; each axis 5 warmups + 10 samples, 32 synthetic Thumb deltas in 4 bursts; times end at ApplicationIdle after layout, not physical pointer-to-display latency.",
            samples
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "resize-benchmark-result.txt"), "PASS: resize measurements completed; see raw samples and counters in resize-benchmark.json.");
    }
}
