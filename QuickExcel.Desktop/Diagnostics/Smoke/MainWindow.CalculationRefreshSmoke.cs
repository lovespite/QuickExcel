using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCalculationRefreshSmokeAsync(Action<bool, string> check)
    {
        async Task Idle() => await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
        var calculator = new RefreshSmokeCalculator(_services.Calculator);
        InstallDocument(WorkbookDocument.CreateBlank(new(_services.Storage, calculator)), null, "刷新回归.xlsq");
        ApplyValues([(0, 0, "2"), (0, 1, "=A1*2")]);
        await RecalculateAsync();
        GridSurface.Select(new(1, 1)); Freeze_Click(new MenuItem(), new()); await Idle();
        GridSurface.Select(new(80, 1)); await Idle();
        var model = _grid.Control.Model;
        var offset = _grid.Control.FirstVisibleRowScrollIndex;
        var current = _grid.Current;
        var editor = _grid.Editor;
        var focus = Keyboard.FocusedElement;
        var disabled = 0;
        void EnabledChanged(object sender, DependencyPropertyChangedEventArgs e) { if (Equals(e.NewValue, false)) disabled++; }
        Toolbar.IsEnabledChanged += EnabledChanged;
        Workspace.IsEnabledChanged += EnabledChanged;
        var suppression = _suppressAutoCalculation;
        try
        {
            var gate = calculator.HoldNext();
            var running = RecalculateAsync();
            try
            {
                await gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
                check(_busy && Toolbar.IsEnabled && Workspace.IsEnabled && SheetGrid.IsEnabled,
                    "Recalculation keeps toolbar and grid enabled without disabled visual transitions");
                check(ReferenceEquals(focus, Keyboard.FocusedElement) && ReferenceEquals(editor, _grid.Editor),
                    "Recalculation keeps the focused editor attached");
                var key = new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(this), 0, Key.A)
                    { RoutedEvent = Keyboard.PreviewKeyDownEvent };
                editor.RaiseEvent(key);
                var text = new TextCompositionEventArgs(Keyboard.PrimaryDevice, new TextComposition(InputManager.Current, editor, "中文"))
                    { RoutedEvent = TextCompositionManager.PreviewTextInputStartEvent };
                editor.RaiseEvent(text);
                var mouse = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left) { RoutedEvent = Mouse.PreviewMouseDownEvent };
                FormulaBar.RaiseEvent(mouse);
                check(key.Handled && text.Handled && mouse.Handled && !_editState.IsComposing && !_editState.GridEditing,
                    "Busy preview gates block keyboard, composition and mouse edits without disabling controls");
            }
            finally { gate.Release.TrySetResult(); }
            check(await running, "Held recalculation commits successfully");
            await Idle();
            check(disabled == 0, "Calculation emits no disabled transitions");
            check(ReferenceEquals(model, _grid.Control.Model) && _grid.Control.FrozenRowCount == 1 && _grid.Control.FrozenColumnCount == 1,
                "Calculation preserves the model and both frozen axes");
            check(current == _grid.Current && _grid.Control.FirstVisibleRowScrollIndex == offset && ReferenceEquals(editor, _grid.Editor),
                "Calculation preserves current cell, scroll offset and prepared editor");

            ApplyValues([(0, 0, "3")]); await RecalculateAsync(); await Idle();
            check(_grid.DisplayText(0, 1) == "6" && _grid.Control.GetModelCellBounds(0, 1) is not null,
                "Changed formula results update frozen rows even when the main row is offscreen");
            GridSurface.SelectRange(new(0, 0, 0, 1), new(0, 0)); await Idle();
            var selection = GridSurface.Selection.ToArray();
            var calls = calculator.Calls;
            _suppressAutoCalculation = false;
            ApplyRibbonStyle(new(Bold: true, FillColor: "FFFFFF00", NumberFormatId: 2)); await Idle();
            var style = _document.Snapshot.Styles[_document.Snapshot.Sheets[0].GetStyleId(0, 1)];
            check(_grid.DisplayText(0, 1) == "6.00" && style.Bold && style.FillColor == "FFFFFF00" && _grid.Artwork.IsLoaded,
                "Formatting projection updates the frozen cell without replacing the model");
            check(calculator.Calls == calls && !_calculationPending && selection.SequenceEqual(GridSurface.Selection),
                "Formatting preserves the range selection and never schedules calculation");
            await UndoAsync(false); await Idle();
            check(_grid.DisplayText(0, 1) == "6" && CurrentView.FrozenRows == 1,
                "Formatting undo restores display and preserves frozen view");
            await UndoAsync(true); await Idle();
            check(calculator.Calls == calls && !_calculationPending && _document.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.State == CalculationState.Fresh,
                "Formatting undo and redo retain fresh results without invoking the calculator");

            _suppressAutoCalculation = true;
            GridSurface.Select(new(0, 0)); FormulaBar.Text = "4";
            ApplyRibbonStyle(new(Italic: true));
            check(_calculationPending && _document.GetCellInput(0, 0, 0) == "4",
                "Formatting still commits an existing content draft and retains its pending recalculation");
            await RecalculateAsync();
            check(_document.GetCellValue(0, 0, 1) == "8.00", "Committed content draft recalculates dependent formulas");

            gate = calculator.HoldNext();
            running = RecalculateAsync();
            try
            {
                await gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
                var stopMouse = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left) { RoutedEvent = Mouse.PreviewMouseDownEvent };
                StopCalculationButton.RaiseEvent(stopMouse);
                check(!stopMouse.Handled, "Busy mouse gate leaves Stop accessible");
                StopCalculation_Click(this, new());
            }
            finally { gate.Release.TrySetResult(); }
            check(!await running && !_busy && _calculationPending && _document.GetCellValue(0, 0, 1) == "8.00",
                "Cancellation releases the busy gate and preserves the last result for retry");
            check(await RecalculateAsync() && !_calculationPending, "Recalculation succeeds after cancellation");
        }
        finally
        {
            _suppressAutoCalculation = suppression;
            Toolbar.IsEnabledChanged -= EnabledChanged;
            Workspace.IsEnabledChanged -= EnabledChanged;
        }
    }

    private sealed class RefreshSmokeCalculator(IWorkbookCalculator inner) : IWorkbookCalculator
    {
        public sealed class Gate
        {
            public TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
            public TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        }
        private Gate? _next;
        public int Calls { get; private set; }
        public bool LastForceFull { get; private set; }
        public CalculationStatistics? LastStatistics { get; private set; }
        public Gate HoldNext() => _next = new();
        public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default)
        {
            Calls++;
            LastForceFull = request.ForceFull;
            if (_next is { } gate)
            {
                _next = null;
                gate.Entered.TrySetResult();
                gate.Release.Task.Wait(cancellationToken);
            }
            var result = inner.Calculate(request, cancellationToken);
            LastStatistics = result.Statistics;
            return result;
        }
    }
}
