using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace QuickExcel.SandboxProbe;

internal static class WfpEvidence
{
    internal static bool Analyze(string directory)
    {
        using var request = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "capture-request.json")));
        using var network = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "network.json")));
        var config = request.RootElement;
        var start = config.GetProperty("startedUtc").GetDateTimeOffset();
        var childPath = config.GetProperty("executable").GetString()!;
        var listenerPath = config.GetProperty("listenerExecutable").GetString()!;
        var filters = XDocument.Load(Path.Combine(directory, "wfp-filters.xml")).Descendants("item")
            .Where(e => e.Element("filterId") is not null).ToDictionary(e => e.Element("filterId")!.Value);
        bool IsPath(XElement header, string path)
        {
            var hex = header.Element("appId")?.Element("data")?.Value;
            if (string.IsNullOrEmpty(hex)) return false;
            var nt = Encoding.Unicode.GetString(Convert.FromHexString(hex)).TrimEnd('\0');
            // netsh uses an NT volume path. Compare the full path below the drive, not just the file name.
            return nt.EndsWith(path[Path.GetPathRoot(path)!.Length..].Insert(0, "\\"), StringComparison.OrdinalIgnoreCase);
        }
        XElement[] Drops(string name, string path, Func<XElement, bool> matches) =>
            XDocument.Load(Path.Combine(directory, name)).Root!.Elements("item").Where(e =>
            {
                var header = e.Element("header")!;
                return DateTimeOffset.Parse(header.Element("timeStamp")!.Value) >= start && IsPath(header, path) &&
                    e.Element("type")?.Value.EndsWith("CLASSIFY_DROP", StringComparison.Ordinal) == true &&
                    filters.TryGetValue(e.Element("classifyDrop")?.Element("filterId")?.Value ?? "", out var filter) &&
                    filter.Element("action")?.Element("type")?.Value == "FWP_ACTION_BLOCK" && matches(e);
            }).ToArray();
        XElement[] Loop(string version) => Drops("wfp-loopback-" + version + ".xml", listenerPath, e =>
            e.Element("classifyDrop")?.Element("isLoopback")?.Value == "true" &&
            e.Element("header")?.Element("localPort")?.Value == config.GetProperty(version + "Port").GetInt32().ToString() &&
            e.Element("header")?.Element("ipProtocol")?.Value == "6" &&
            e.Element("internalFields")?.Element("filterOrigin")?.Value == "AppContainer Loopback");
        var ipv4 = Loop("IPv4"); var ipv6 = Loop("IPv6");
        var dns = Drops("wfp-netevents.xml", childPath, e =>
            e.Element("header")?.Element("remotePort")?.Value == "53" &&
            e.Element("header")?.Element("ipProtocol")?.Value == "17" &&
            e.Element("internalFields")?.Element("filterOrigin")?.Value == "UWP Default Outbound Block Rule");
        var controls = network.RootElement.GetProperty("positiveControl");
        var isolated = network.RootElement.GetProperty("productionLauncher");
        string[] probes = ["loopbackIPv4", "loopbackIPv6", "externalTcp", "systemDns", "nativeDnsWireOnly", "directDnsUdp", "directDnsTcp"];
        var positive = probes.All(p => controls.GetProperty("checks").GetProperty(p).GetProperty("Outcome").GetString() == "allowed") &&
            controls.GetProperty("listenerAcceptedIPv4").GetBoolean() && controls.GetProperty("listenerAcceptedIPv6").GetBoolean();
        var noAccessObserved = probes.All(p => isolated.GetProperty("checks").GetProperty(p).GetProperty("Outcome").GetString() is "denied" or "inconclusive:timeout") &&
            !isolated.GetProperty("listenerAcceptedIPv4").GetBoolean() && !isolated.GetProperty("listenerAcceptedIPv6").GetBoolean();
        var verified = positive && noAccessObserved && ipv4.Length > 0 && ipv6.Length > 0 && dns.Length > 0;
        var summary = new
        {
            scope = "This Windows machine, these endpoints, production launcher with development syscall payload; not a release security gate.",
            positiveControlsPassed = positive, noAccessObserved, loopbackIPv4PolicyDropObserved = ipv4.Length > 0,
            loopbackIPv6PolicyDropObserved = ipv6.Length > 0, dnsUdpPolicyDropObserved = dns.Length > 0,
            localNetworkEvidenceValidated = verified, releaseIsolationValidated = false,
            drops = ipv4.Concat(ipv6).Concat(dns).Select(e => new
            {
                timestamp = e.Element("header")!.Element("timeStamp")!.Value,
                protocol = e.Element("header")!.Element("ipProtocol")!.Value,
                localPort = e.Element("header")!.Element("localPort")!.Value,
                remotePort = e.Element("header")!.Element("remotePort")!.Value,
                filterId = e.Element("classifyDrop")!.Element("filterId")!.Value,
                filterOrigin = e.Element("internalFields")!.Element("filterOrigin")!.Value
            }).ToArray()
        };
        var file = Path.Combine(directory, "wfp-summary.json");
        File.WriteAllText(file, JsonSerializer.Serialize(summary, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(file);
        return verified;
    }
}
