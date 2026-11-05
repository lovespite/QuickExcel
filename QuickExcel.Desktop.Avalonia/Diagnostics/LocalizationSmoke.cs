using System.Globalization;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunLocalizationSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>();
        Exception? failure = null;
        void Check(bool value, string label)
        {
            if (!value) throw new InvalidOperationException(label);
            checks.Add(label);
        }
        async Task CheckPage(Web.LocalWebView browser, string expression, string label)
        {
            var result = await browser.View.InvokeScript(expression);
            Check(result == "true", label + " (browser returned " + result + ")");
        }
        try
        {
            var english = UiText.Language == UiText.English;
            Check(UiText.Get("设置") == (english ? "Settings" : "设置"), "Catalog loaded before the window");
            var value = "设置 {0} 😀";
            Check(UiText.Format($"已执行：{value}") == (english ? "Completed: " : "已执行：") + value,
                "Formatting preserves user data and braces");
            Check(UiText.Format($"计数：{1234:N0}").EndsWith(1234.ToString("N0", CultureInfo.CurrentCulture)),
                "UI language preserves the user's numeric formatting culture");
            Check(UiText.Get("unregistered-key") == "unregistered-key", "Missing keys retain their source");
            _workspace.New(); ShowBook();
            Document.SetCell(0, 0, 0, value);
            var report = TypedWorkbookStorage.InspectCompatibility(Document.Snapshot);
            Check(report.Diagnostics.Any(d => d.Code == "workbook.modeled-subset" &&
                d.Message.StartsWith(english ? "Supported cell types" : "单元格类型")), "Compatibility details use the chosen language");
            try
            {
                _workspace.Editor.Execute(new CopySheetCommand(Document.Snapshot.ActiveSheetId, Document.Snapshot.Sheets[0].Name, 1));
                throw new InvalidOperationException("Duplicate sheet was accepted");
            }
            catch (ArgumentException error)
            {
                Check(error.Message == (english ? "A worksheet with this name already exists." : "工作表名称已存在。"),
                    "Core validation uses the host catalog");
            }
            const string source = "// 设置 {0}\nhost.functions.register('LANGUAGE', () => 1);";
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.js", source) }, ["main"], scriptApiVersion: 3)));
            await OpenScriptsAsync();
            await CheckPage(_scriptPage!.Browser, english
                ? "document.documentElement.lang === 'en-US' && document.querySelector('[data-action]').textContent === 'Suggest'"
                : "document.documentElement.lang === 'zh-CN' && document.querySelector('[data-action]').textContent === '补全'",
                "The native Monaco page follows the stored language");
            Check(await CommitScriptsAsync() && Document.Snapshot.ScriptModules.Modules["main"].Source == source,
                "Localization preserves script source");
            await CloseScriptWindowAsync();
            await OpenPrintPreviewAsync();
            await CheckPage(_pdfPage!.Browser, english
                ? "document.documentElement.lang === 'en-US' && document.getElementById('export').textContent === 'Export PDF'"
                : "document.documentElement.lang === 'zh-CN' && document.getElementById('export').textContent === '导出 PDF'",
                "The native PDF page follows the stored language");
            Check(Document.GetCellInput(0, 0, 0) == value, "Localization and preview preserve workbook text");
            await ClosePrintPreviewAsync();
        }
        catch (Exception error) { failure = error; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(
            new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true;
        (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
