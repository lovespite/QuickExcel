using System.Diagnostics;
using System.ComponentModel;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.Json;
using QuickExcel.Runtime;
using QuickExcel.Scripting;
using QuickExcel.SandboxProbe;

// Development-only OS syscall probes. This executable is never included in a release.
if (args.Length == 2 && args[0] == "--analyze-wfp") return WfpEvidence.Analyze(args[1]) ? 0 : 2;
if (args.Length == 2 && args[0] == "--network") return await NetworkValidation.Run(args[1]);
if (args.Length == 2 && args[0] == "--network-wfp") return await NetworkValidation.Run(args[1], captureWfp: true);
if (args.Length == 0 && Console.IsInputRedirected)
{
    var config = JsonSerializer.Deserialize<NetworkConfiguration>((await Console.In.ReadLineAsync())!)!;
    Console.WriteLine(JsonSerializer.Serialize(await NetworkChecks.Run(config.IPv4Port, config.IPv6Port, config.DnsServer)));
    return 0;
}
if (args.FirstOrDefault() == "--noop") return 0;
if (args.FirstOrDefault() == "--child")
{
    var results = new Dictionary<string, string>();
    void Probe(string name, Action action)
    {
        try { action(); results[name] = "allowed"; }
        catch (Exception e)
        {
            results[name] = e is TimeoutException ? "inconclusive:timeout"
                : e is SocketException socket ? $"failed:{socket.SocketErrorCode}"
                : e is Win32Exception win32 ? $"failed:Win32:{win32.NativeErrorCode}"
                : $"failed:{e.GetType().Name}:{e.HResult:X8}";
        }
    }
    results["appContainer"] = Native.TokenFlag(29).ToString();
    results["restricted"] = Native.TokenFlag(21).ToString();
    Probe("nativeRuntime", () =>
    {
        using var session = new QuickJsSession(new Dictionary<string, string>
        {
            ["main"] = "import {x} from 'helper';quickExcel.functions.register('F',()=>x+1);",
            ["helper"] = "export const x=41;"
        }, ["main"]);
        if (!Equals(session.Invoke("F", []), 42d)) throw new InvalidOperationException();
    });
    Probe("privateRead", () => File.ReadAllText(args[1]));
    Probe("privateWrite", () => File.WriteAllText(Path.Combine(Path.GetDirectoryName(args[1])!, "child-write.txt"), "synthetic probe"));
    Probe("payloadWrite", () => File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "child-write.txt"), "synthetic probe"));
    Probe("inheritedCanaryHandle", () =>
    {
        var bytes = new byte[64];
        if (!Native.ReadFile(new IntPtr(long.Parse(args[3])), bytes, bytes.Length, out var count, IntPtr.Zero) ||
            System.Text.Encoding.UTF8.GetString(bytes, 0, count) != "synthetic QuickExcel canary")
            throw new UnauthorizedAccessException();
    });
    Probe("loopback", () =>
    {
        using var socket = new TcpClient();
        socket.ConnectAsync(IPAddress.Loopback, int.Parse(args[2])).WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult();
    });
    // No workbook or user data is sent. DNS lookup and TCP handshake only.
    Probe("dns", () => Dns.GetHostAddressesAsync("example.com").WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult());
    Probe("externalTcp", () =>
    {
        using var socket = new TcpClient();
        socket.ConnectAsync("1.1.1.1", 443).WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult();
    });
    Probe("childProcess", () =>
    {
        using var child = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--noop") { UseShellExecute = false, CreateNoWindow = true })!;
        if (!child.WaitForExit(2000)) { child.Kill(); throw new TimeoutException(); }
    });
    Console.WriteLine(JsonSerializer.Serialize(results));
    return 0;
}

if (args.Length != 1 || !File.Exists(Path.Combine(args[0], "QuickExcel.slnx")))
    throw new ArgumentException("Pass the QuickExcel repository root. All probe files stay in its artifacts directory.");
var root = Path.GetFullPath(args[0]);
var run = Path.Combine(root, "artifacts", "sandbox-probe", DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N")[..8]);
var payload = Path.Combine(run, "payload");
var privateFolder = Path.Combine(run, "private");
Directory.CreateDirectory(payload); Directory.CreateDirectory(privateFolder);
foreach (var path in Directory.EnumerateFiles(AppContext.BaseDirectory)) File.Copy(path, Path.Combine(payload, Path.GetFileName(path)));
var canary = Path.Combine(privateFolder, "canary.txt");
File.WriteAllText(canary, "synthetic QuickExcel canary");
var privateAcl = new DirectorySecurity();
privateAcl.SetAccessRuleProtection(true, false);
privateAcl.AddAccessRule(new(WindowsIdentity.GetCurrent().User!, FileSystemRights.FullControl,
    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
new DirectoryInfo(privateFolder).SetAccessControl(privateAcl);
var report = new Dictionary<string, object>
{
    ["os"] = Environment.OSVersion.VersionString, ["runtime"] = RuntimeInformation.FrameworkDescription,
    ["releaseIsolationValidated"] = false
};
using var listener = new TcpListener(IPAddress.Loopback, 0);
listener.Start();
var port = ((IPEndPoint)listener.LocalEndpoint).Port;
using var leak = Native.OpenInheritableFile(canary);
var profile = "QuickExcel.Probe." + Guid.NewGuid().ToString("N");
IntPtr sid = IntPtr.Zero;
var originalAcl = new DirectoryInfo(payload).GetAccessControl();
var profileCreated = false;
var cleanup = true;
try
{
    var hr = Native.CreateAppContainerProfile(profile, profile, "Temporary QuickExcel development probe", IntPtr.Zero, 0, out sid);
    if (hr != 0) Marshal.ThrowExceptionForHR(hr);
    profileCreated = true;
    var acl = new DirectoryInfo(payload).GetAccessControl();
    acl.AddAccessRule(new(new SecurityIdentifier(sid), FileSystemRights.ReadAndExecute,
        InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
    new DirectoryInfo(payload).SetAccessControl(acl);
    foreach (var policy in new[] { "handle-control", "job-only", "stripped-token", "appcontainer" })
    {
        try
        {
            Native.Rewind(leak);
            report[policy] = await Native.Run(Path.Combine(payload, "QuickExcel.SandboxProbe.exe"),
                $"--child \"{canary}\" {port} {leak.DangerousGetHandle().ToInt64()}", policy, sid, leak.DangerousGetHandle());
        }
        catch (Exception e) { report[policy] = new { launchFailed = e.GetType().Name, code = e.HResult.ToString("X8"), detail = e.Message }; }
    }
}
catch (Exception e) { report["setupFailure"] = new { error = e.GetType().Name, code = e.HResult.ToString("X8"), detail = e.Message }; }
finally
{
    try { new DirectoryInfo(payload).SetAccessControl(originalAcl); }
    catch (Exception e) { report["aclCleanupFailure"] = e.GetType().Name; cleanup = false; }
    if (sid != IntPtr.Zero) Native.FreeSid(sid);
    if (profileCreated) { var hr = Native.DeleteAppContainerProfile(profile); report["profileCleanupHresult"] = hr; cleanup &= hr == 0; }
    report["cleanupSucceeded"] = cleanup;
}
var reportPath = Path.Combine(run, "probe.json");
var completed = cleanup && new[] { "handle-control", "job-only", "stripped-token", "appcontainer" }
    .All(policy => report.GetValueOrDefault(policy) is Dictionary<string, string>);
report["probeCompleted"] = completed;
File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(reportPath);
Console.WriteLine(JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
// Exit zero means the probes completed, not that an OS security gate passed.
return completed ? 0 : 1;
