using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckEventTemplateInsertionAsync(Func<object, Task<JsonElement>> test, Func<string, Task> select,
        Action<bool, string> check, string directory)
    {
        await select("action"); check(_scriptEventButton?.IsVisible == false, "The event insertion button is hidden for Actions");
        await select("events"); check(_scriptEventButton?.IsVisible == true, "Selecting an Events module shows Add Event Handler");
        const string source = "// before\n\n// after\n";
        await test(new Dictionary<string, object?> { ["action"] = "edit", ["module"] = "events", ["text"] = source });
        await test(new Dictionary<string, object?> { ["action"] = "assistance", ["offset"] = "// before\n".Length });
        _dialogDriver = async window =>
        {
            var dialog = (ScriptEventDialog)window;
            dialog.Search.Text = "cell.click";
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            check(dialog.Events.ItemCount == 1 && dialog.Description.Text!.Contains("点击") && dialog.Timing.Text!.Contains("网格"),
                "Searching selects the event and shows its Chinese description and trigger timing");
            Capture(dialog, "script-event-dialog-zh.png");
            dialog.Confirm.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        };
        await AddEventHandlerAsync();
        var inserted = (await test(new Dictionary<string, object?> { ["action"] = "inspect" })).GetProperty("source").GetString()!;
        check(inserted.IndexOf("// before", StringComparison.Ordinal) < inserted.IndexOf("export function onCellClick(context: QuickExcel.EventContext): void", StringComparison.Ordinal) &&
            inserted.IndexOf("export function onCellClick", StringComparison.Ordinal) < inserted.IndexOf("// after", StringComparison.Ordinal),
            "The Events template is inserted at the captured cursor in the current TypeScript model");
        check(!inserted.Contains(".on("), "The inserted handler has an exported signature without manual registration");
        await _scriptPage!.ExecuteEditorCommandAsync("undo");
        check((await test(new Dictionary<string, object?> { ["action"] = "inspect" })).GetProperty("source").GetString()!.Replace("\r\n", "\n") == source,
            "One editor undo removes only the inserted template");
        var anchor = await _scriptPage!.CaptureInsertionAsync("events");
        await test(new Dictionary<string, object?> { ["action"] = "assistance", ["typeText"] = "// changed\n" });
        var rejected = await _scriptPage.Browser.RequestAsync("insertTemplate", "templateInserted", new Dictionary<string, object?> { ["anchor"] = anchor, ["text"] = "must-not-appear" });
        check(!rejected.Payload.GetProperty("accepted").GetBoolean() && !_scriptPage.Drafts.Failed,
            "A stale cursor anchor rejects insertion without losing or failing the draft");
        var beforeCancel = (await test(new Dictionary<string, object?> { ["action"] = "inspect" })).GetProperty("source").GetString();
        var priorLanguage = UiText.Language;
        try
        {
            UiText.Initialize(UiText.English);
            _dialogDriver = async window =>
            {
                var dialog = (ScriptEventDialog)window; dialog.Search.Text = "cell.click";
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                check(dialog.Title == "Add Event Handler" && dialog.Description.Text!.Contains("click") && dialog.Timing.Text!.Contains("grid"),
                    "The same event dialog uses English descriptions and timing under the English UI language");
                Capture(dialog, "script-event-dialog-en.png");
                dialog.Search.Text = "no matching event";
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                check(dialog.Events.ItemCount == 0 && !dialog.Confirm.IsEnabled, "An empty event search disables confirmation");
                window.Close(null);
            };
            await AddEventHandlerAsync();
        }
        finally { UiText.Initialize(priorLanguage); }
        check((await test(new Dictionary<string, object?> { ["action"] = "inspect" })).GetProperty("source").GetString() == beforeCancel,
            "Cancelling event selection preserves the source and cursor-owned draft");
        await test(new Dictionary<string, object?> { ["action"] = "edit", ["module"] = "events", ["text"] = "export function onCellClick(context: QuickExcel.EventContext): void { console.log(context.event.name); }" });

        void Capture(Window window, string name)
        {
            using var image = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(window.Bounds.Width), (int)Math.Ceiling(window.Bounds.Height)), new Vector(96, 96));
            image.Render(window); image.Save(Path.Combine(directory, name), PngBitmapEncoderOptions.Default);
        }
    }
}
