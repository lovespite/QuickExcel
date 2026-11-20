using System.Text.Json.Serialization;

namespace QuickExcel.Core;

/// <summary>User preferences for host-brokered Actions and session network diagnostics.</summary>
public sealed record ScriptNetworkOptions(int BodyLimitMiB = 8, int ActionTimeoutSeconds = 10,
    int PanelMemoryLimitMiB = 50, int PanelEntryLimit = 200, bool AllowNetwork = true)
{
    public const int MinBodyLimitMiB = 1, MaxBodyLimitMiB = 10;
    public const int MinActionTimeoutSeconds = 1, MaxActionTimeoutSeconds = 300;
    public const int MinPanelMemoryLimitMiB = 1, MaxPanelMemoryLimitMiB = 512;
    public const int MinPanelEntryLimit = 1, MaxPanelEntryLimit = 5000;
    [JsonIgnore] public bool IsValid => BodyLimitMiB is >= MinBodyLimitMiB and <= MaxBodyLimitMiB &&
        ActionTimeoutSeconds is >= MinActionTimeoutSeconds and <= MaxActionTimeoutSeconds &&
        PanelMemoryLimitMiB is >= MinPanelMemoryLimitMiB and <= MaxPanelMemoryLimitMiB &&
        PanelEntryLimit is >= MinPanelEntryLimit and <= MaxPanelEntryLimit;
    [JsonIgnore] public int BodyLimitBytes => checked(BodyLimitMiB * 1024 * 1024);
    [JsonIgnore] public TimeSpan ActionTimeout => TimeSpan.FromSeconds(ActionTimeoutSeconds);
    [JsonIgnore] public long PanelMemoryLimitBytes => (long)PanelMemoryLimitMiB * 1024 * 1024;
    public void Validate()
    {
        if (!IsValid) throw new ArgumentOutOfRangeException(nameof(ScriptNetworkOptions), "网络设置超出允许范围。");
    }
}
