using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using QuickExcel.Core;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

internal sealed partial class SettingsWindow
{
    private void NetworkSettings()
    {
        _page.Children.Add(new Separator());
        _page.Children.Add(Heading(UiText.Get("脚本网络与诊断限额"), 18));
        var allowNetwork = new CheckBox { Name = "AllowScriptNetwork", Content = UiText.Get("允许访问网络"), IsChecked = _settings.Value.ScriptNetwork.AllowNetwork };
        AutomationProperties.SetName(allowNetwork, UiText.Get("允许访问网络"));
        allowNetwork.IsCheckedChanged += (_, _) =>
        {
            Save(_settings.Value with { ScriptNetwork = _settings.Value.ScriptNetwork with { AllowNetwork = allowNetwork.IsChecked == true } });
            if (_settings.Error is not null) allowNetwork.IsChecked = _settings.Value.ScriptNetwork.AllowNetwork;
        };
        _page.Children.Add(allowNetwork);
        _page.Children.Add(Note(UiText.Get("关闭后立即禁止所有脚本网络请求，网络权限申请也会直接拒绝。重新打开后需重新授权。")));
        var options = _settings.Value.ScriptNetwork;
        var rows = new StackPanel { Spacing = 10 };
        NumericUpDown Limit(string name, string label, int value, int minimum, int maximum)
        {
            var row = new LayoutGrid { ColumnDefinitions = new("*,130"), ColumnSpacing = 14 };
            var caption = new TextBlock { Text = UiText.Get(label), TextWrapping = global::Avalonia.Media.TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };
            row.Children.Add(caption);
            var input = new NumericUpDown { Name = name, Minimum = minimum, Maximum = maximum, Increment = 1, FormatString = "0", Value = value,
                HorizontalAlignment = HorizontalAlignment.Stretch };
            AutomationProperties.SetName(input, UiText.Get(label));
            ToolTip.SetTip(input, UiText.Format($"整数范围：{minimum}–{maximum}"));
            LayoutGrid.SetColumn(input, 1); row.Children.Add(input); rows.Children.Add(row);
            return input;
        }
        var body = Limit("NetworkBodyLimitMiB", "单次请求 / 响应大小上限（MiB）", options.BodyLimitMiB,
            ScriptNetworkOptions.MinBodyLimitMiB, ScriptNetworkOptions.MaxBodyLimitMiB);
        var timeout = Limit("NetworkActionTimeoutSeconds", "Action 执行超时（秒，含网络等待）", options.ActionTimeoutSeconds,
            ScriptNetworkOptions.MinActionTimeoutSeconds, ScriptNetworkOptions.MaxActionTimeoutSeconds);
        var entries = Limit("NetworkPanelEntryLimit", "网络面板记录条数上限", options.PanelEntryLimit,
            ScriptNetworkOptions.MinPanelEntryLimit, ScriptNetworkOptions.MaxPanelEntryLimit);
        var memory = Limit("NetworkPanelMemoryLimitMiB", "网络面板内存上限（MiB）", options.PanelMemoryLimitMiB,
            ScriptNetworkOptions.MinPanelMemoryLimitMiB, ScriptNetworkOptions.MaxPanelMemoryLimitMiB);
        _page.Children.Add(rows);
        _page.Children.Add(Note(UiText.Get("默认：请求和响应分别 8 MiB，Action 总超时 10 秒，网络面板最多 200 条记录、50 MiB 内存。执行限额用于下一次 Action；调低面板限额会立即释放最旧的记录或内容。")));
        var apply = new Button { Name = "ApplyNetworkLimits", Content = UiText.Get("应用网络限额") };
        apply.Click += (_, _) =>
        {
            var inputs = new[] { body, timeout, entries, memory };
            if (inputs.Any(input => input.Value is not { } value || value != decimal.Truncate(value) || value < input.Minimum || value > input.Maximum))
            { _feedback.Text = UiText.Get("请输入各项允许范围内的整数。"); return; }
            var updated = new ScriptNetworkOptions((int)body.Value!.Value, (int)timeout.Value!.Value, (int)memory.Value!.Value, (int)entries.Value!.Value,
                _settings.Value.ScriptNetwork.AllowNetwork);
            Save(_settings.Value with { ScriptNetwork = updated });
        };
        _page.Children.Add(apply);
    }
}
