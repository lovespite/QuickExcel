using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Markdown.Avalonia;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunLicenseSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>(); Exception? failure = null;
        SettingsWindow? window = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control root, string name) where T : Control => root.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        async Task Capture(string name)
        {
            await Layout();
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)window!.Bounds.Width, (int)window.Bounds.Height), new Vector(96, 96));
            bitmap.Render(window); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            window = new SettingsWindow(_settings, _recentFiles, SettingsSection.About);
            window.Show(this); await Layout();
            var selector = Named<ComboBox>(window, "SettingsLicenseSelector");
            var markdown = Named<MarkdownScrollViewer>(window, "SettingsLicenseMarkdown");
            var text = Named<TextBox>(window, "SettingsLicenseText");
            var readingArea = Named<Control>(window, "SettingsLicenseViewer");
            var pagination = Named<StackPanel>(window, "SettingsLicensePagination");
            var fullNotice = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, (string)selector.SelectedItem!));
            Check(markdown.IsVisible && !text.IsVisible && !pagination.IsVisible && markdown.Markdown == fullNotice,
                "The complete Markdown notice is rendered without character-based pagination");
            Check(markdown.GetVisualDescendants().OfType<Control>().Any(c => c.Classes.Contains(global::Markdown.Avalonia.Markdown.Heading1Class)) &&
                markdown.GetVisualDescendants().OfType<Control>().Any(c => c.Classes.Contains(global::Markdown.Avalonia.Markdown.TableClass)),
                "The notice creates native heading and table controls");
            Check(fullNotice.Contains("Markdown.Avalonia.Tight") && fullNotice.Contains("ColorDocument.Avalonia"),
                "Settings includes attribution for the Markdown renderer and its dependencies");
            var normal = readingArea.Bounds.Size;
            await Capture("licenses-default");
            window.Width = 1180; window.Height = 820; await Layout();
            Check(readingArea.Bounds.Width > normal.Width + 300 && readingArea.Bounds.Height > normal.Height + 150,
                "The license viewport grows in both directions with the window");
            await Capture("licenses-expanded");
            window.Width = window.MinWidth; window.Height = window.MinHeight; await Layout();
            Check(readingArea.Bounds.Height >= 80 && readingArea.Bounds.Width >= 300 && readingArea.Bounds.Height < normal.Height,
                "The minimum window retains a usable license viewport");
            Check(selector.TranslatePoint(new Point(selector.Bounds.Width, selector.Bounds.Height), window) is { } corner &&
                corner.X <= window.Bounds.Width && corner.Y <= window.Bounds.Height,
                "The license selector stays inside the minimum window");
            await Capture("licenses-minimum");
            var markdownScroll = markdown.GetVisualDescendants().OfType<ScrollViewer>().First();
            markdownScroll.ScrollToEnd(); await Layout();
            Check(markdownScroll.Offset.Y > 0, "The entire Markdown document remains scrollable at the minimum size");
            selector.SelectedItem = selector.Items.OfType<string>().Single(p => p.EndsWith("OFFICE-SCRIPTS-NOTICE.md")); await Layout();
            Check(markdown.Markdown!.Contains("CC BY 4.0") && markdown.ScrollValue == default,
                "Switching Markdown documents preserves attribution and resets the scroll position");
            selector.SelectedItem = selector.Items.OfType<string>().Single(p => p.EndsWith("OPENMCDF-LICENSE.txt")); await Layout();
            var plainLicense = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, (string)selector.SelectedItem!));
            Check(text.IsVisible && !markdown.IsVisible && pagination.IsVisible && text.Text == plainLicense[..4000],
                "Plain-text licenses retain their exact original content and pagination");
            Check(readingArea.Bounds.Height >= 80, "Plain-text pagination leaves a usable viewport at the minimum size");
            var next = pagination.Children.OfType<Button>().Single(b => Equals(b.Content, UiText.Get("下一页")));
            var previous = pagination.Children.OfType<Button>().Single(b => Equals(b.Content, UiText.Get("上一页")));
            var reconstructed = text.Text;
            while (next.IsEnabled) { next.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); reconstructed += text.Text; }
            Check(reconstructed == plainLicense && previous.IsEnabled, "All plain-text pages reconstruct the license without omissions");
            previous.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Check(next.IsEnabled && text.CaretIndex == 0, "Previous-page navigation remains available and resets the caret");
            await Capture("licenses-plain-minimum");
            selector.SelectedItem = selector.Items.OfType<string>().Single(p => p.EndsWith("MARKDOWN-AVALONIA-LICENSE.txt")); await Layout();
            Check(text.Text!.Contains("2010 Bevan Arps, 2020 Whistyun") && text.Text.Contains("Permission is hereby granted"),
                "The upstream MIT copyright and license are distributed and readable offline");
            var navigation = Named<ListBox>(window, "SettingsNavigation");
            navigation.SelectedItem = UiText.Get("历史管理"); await Layout();
            Check(!window.GetVisualDescendants().OfType<Control>().Any(c => c.Name == "SettingsLicenseViewer"),
                "Changing settings sections removes the license reading area");
            navigation.SelectedItem = UiText.Get("关于与许可"); await Layout();
            Check(Named<MarkdownScrollViewer>(window, "SettingsLicenseMarkdown").Markdown == fullNotice,
                "Reopening About restores the initial rendered notice");
        }
        catch (Exception ex) { failure = ex; if (window is not null) await Capture("licenses-failure"); }
        finally { window?.Close(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
