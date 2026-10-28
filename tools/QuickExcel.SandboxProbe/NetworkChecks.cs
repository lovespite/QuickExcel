using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace QuickExcel.SandboxProbe;

internal record NetworkResult(string Outcome, long Milliseconds, int? NativeCode = null, string? Detail = null);

internal static class NetworkChecks
{
    internal static async Task<Dictionary<string, NetworkResult>> Run(int ipv4Port, int ipv6Port, string dnsServer)
    {
        var probes = new Dictionary<string, Func<Task>>
        {
            ["loopbackIPv4"] = () => Tcp(IPAddress.Loopback, ipv4Port),
            ["loopbackIPv6"] = () => Tcp(IPAddress.IPv6Loopback, ipv6Port),
            ["externalTcp"] = () => Tcp(IPAddress.Parse("1.1.1.1"), 443),
            ["systemDns"] = async () => { _ = await Dns.GetHostAddressesAsync("example.com."); },
            ["nativeDnsWireOnly"] = () => Task.Run(() => NativeDns(0x100 | 0x8 | 0x40 | 0x80 | 0x800 | 0x1000)),
            ["nativeDnsCacheOnly"] = () => Task.Run(() => NativeDns(0x10 | 0x1000)),
            ["directDnsUdp"] = () => DirectDns(dnsServer, tcp: false),
            ["directDnsTcp"] = () => DirectDns(dnsServer, tcp: true)
        };
        var results = await Task.WhenAll(probes.Select(async pair =>
        {
            var clock = Stopwatch.StartNew();
            try { await pair.Value().WaitAsync(TimeSpan.FromSeconds(12)); return KeyValuePair.Create(pair.Key, new NetworkResult("allowed", clock.ElapsedMilliseconds)); }
            catch (TimeoutException) { return KeyValuePair.Create(pair.Key, new NetworkResult("inconclusive:timeout", clock.ElapsedMilliseconds)); }
            catch (SocketException e) { return KeyValuePair.Create(pair.Key, new NetworkResult(e.SocketErrorCode == SocketError.AccessDenied ? "denied" : "inconclusive:socket", clock.ElapsedMilliseconds, e.NativeErrorCode, e.SocketErrorCode.ToString())); }
            catch (DnsException e) { return KeyValuePair.Create(pair.Key, new NetworkResult(e.Code is 5 or 10013 ? "denied" : "inconclusive:dns", clock.ElapsedMilliseconds, e.Code)); }
            catch (Exception e) { return KeyValuePair.Create(pair.Key, new NetworkResult("inconclusive:error", clock.ElapsedMilliseconds, e.HResult, e.GetType().Name)); }
        }));
        var report = results.ToDictionary();
        // This diagnoses missing capabilities. It is supplementary evidence, not a packet-drop event.
        foreach (var address in new[] { "127.0.0.1", "::1", dnsServer, "1.1.1.1" }.Distinct())
        {
            var code = NetworkIsolationDiagnoseConnectFailureAndGetInfo(address, out var reason);
            report["capabilityDiagnostic:" + address] = new(code == 0 ? "diagnostic" : "inconclusive:diagnostic", 0, (int)code, "missingCapability=" + reason);
        }
        return report;
    }

    private static async Task Tcp(IPAddress address, int port)
    {
        using var socket = new Socket(address.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(11));
        try { await socket.ConnectAsync(new IPEndPoint(address, port), deadline.Token); }
        catch (OperationCanceledException) { throw new TimeoutException(); }
    }

    private static void NativeDns(uint options)
    {
        var status = DnsQuery("example.com.", 1, options, IntPtr.Zero, out var records, IntPtr.Zero);
        if (records != IntPtr.Zero) DnsRecordListFree(records, 1);
        if (status != 0) throw new DnsException(status);
    }

    private static async Task DirectDns(string server, bool tcp)
    {
        var ip = IPAddress.Parse(server);
        using var socket = new Socket(ip.AddressFamily, tcp ? SocketType.Stream : SocketType.Dgram, tcp ? ProtocolType.Tcp : ProtocolType.Udp);
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(11));
        // Public, fixed query only. Never encode workbook contents, paths or arbitrary user data in DNS.
        byte[] query = [0x51, 0x58, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 7, 101, 120, 97, 109, 112, 108, 101, 3, 99, 111, 109, 0, 0, 1, 0, 1];
        try
        {
            await socket.ConnectAsync(new IPEndPoint(ip, 53), deadline.Token);
            byte[] reply;
            if (tcp)
            {
                using var stream = new NetworkStream(socket, ownsSocket: false);
                await stream.WriteAsync(new byte[] { 0, (byte)query.Length }, deadline.Token);
                await stream.WriteAsync(query, deadline.Token);
                var prefix = new byte[2]; await stream.ReadExactlyAsync(prefix, deadline.Token);
                var size = (prefix[0] << 8) | prefix[1];
                if (size is < 12 or > 4096) throw new IOException("Invalid DNS frame.");
                reply = new byte[size]; await stream.ReadExactlyAsync(reply, deadline.Token);
            }
            else
            {
                await socket.SendAsync(query, SocketFlags.None, deadline.Token);
                var bytes = new byte[4096]; var count = await socket.ReceiveAsync(bytes, SocketFlags.None, deadline.Token);
                reply = bytes[..count];
            }
            if (reply.Length < 12 || reply[0] != query[0] || reply[1] != query[1] || (reply[2] & 0x80) == 0)
                throw new IOException("Invalid DNS response.");
            // Even an NXDOMAIN/REFUSED reply proves network access; do not call it an isolation denial.
        }
        catch (OperationCanceledException) { throw new TimeoutException(); }
    }
    private sealed class DnsException(int code) : Exception { internal int Code => code; }
    [DllImport("dnsapi.dll", CharSet = CharSet.Unicode, EntryPoint = "DnsQuery_W")] private static extern int DnsQuery(string name, ushort type, uint options, IntPtr extra, out IntPtr records, IntPtr reserved);
    [DllImport("dnsapi.dll")] private static extern void DnsRecordListFree(IntPtr records, int freeType);
    [DllImport("Firewallapi.dll", CharSet = CharSet.Unicode)] private static extern uint NetworkIsolationDiagnoseConnectFailureAndGetInfo(string server, out int reason);
}
