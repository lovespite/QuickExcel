using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text.Json;
using QuickExcel.Runtime;

namespace QuickExcel.SandboxProbe;

internal record NetworkConfiguration(int IPv4Port, int IPv6Port, string DnsServer);

internal static class NetworkValidation
{
    internal static async Task<int> Run(string root, bool captureWfp = false)
    {
        if (!File.Exists(Path.Combine(root, "QuickExcel.slnx"))) throw new ArgumentException("Pass the QuickExcel repository root.");
        var directory = Path.Combine(Path.GetFullPath(root), "artifacts", "sandbox-probe", DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(directory);
        // Optional administrator helper reads WFP only. The probes themselves still run under the original user token.
        Process? collector = null;
        var dns = NetworkInterface.GetAllNetworkInterfaces().Where(n => n.OperationalStatus == OperationalStatus.Up)
            .SelectMany(n => n.GetIPProperties().DnsAddresses).FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(a))?.ToString() ?? "1.1.1.1";
        var report = new Dictionary<string, object> { ["os"] = Environment.OSVersion.ToString(), ["runtime"] = Environment.Version.ToString(),
            ["releaseIsolationValidated"] = false, ["dnsServer"] = dns, ["utc"] = DateTime.UtcNow };
        var complete = true;
        if (captureWfp)
        {
            try
            {
                collector = Process.Start(new ProcessStartInfo("powershell.exe")
                {
                    UseShellExecute = true, Verb = "runas", WindowStyle = ProcessWindowStyle.Hidden,
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(root, "tools", "QuickExcel.SandboxProbe", "Collect-Wfp.ps1")}\" -EvidenceDirectory \"{directory}\""
                }) ?? throw new IOException("WFP helper did not start.");
            }
            catch (Exception e) { complete = false; report["wfpCollectorFailure"] = e.ToString(); }
        }
        try
        {
            using var real = (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory().Create("quickExcel.functions.register('CHECK',()=>42);");
            report["actualWorker"] = new { appContainerTokenVerifiedBeforeResume = true, pid = real.ProcessId, udf = real.Invoke("CHECK", []) };
        }
        catch (Exception e) { complete = false; report["actualWorkerFailure"] = e.ToString(); }
        foreach (var isolated in new[] { false, true })
        {
            try
            {
                using var listener4 = new TcpListener(IPAddress.Loopback, 0);
                using var listener6 = new TcpListener(IPAddress.IPv6Loopback, 0);
                listener4.Start(); listener6.Start();
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                async Task<bool> Accept(TcpListener listener)
                {
                    try { using var accepted = await listener.AcceptTcpClientAsync(deadline.Token); return true; }
                    catch (OperationCanceledException) { return false; }
                }
                var accept4 = Accept(listener4); var accept6 = Accept(listener6);
                var config = new NetworkConfiguration(((IPEndPoint)listener4.LocalEndpoint).Port, ((IPEndPoint)listener6.LocalEndpoint).Port, dns);
                Dictionary<string, NetworkResult> checks;
                string? staging = null;
                if (isolated)
                {
                    // Exact production launcher, profile/ACL/env/handle policy and Job; only the test payload differs.
                    using (var child = WorkerProcess.Start(Environment.ProcessPath!, deadline.Token))
                    {
                        staging = child.PayloadDirectory;
                        checks = await Exchange(child.Input, child.Output, child.Error, config, deadline.Token);
                        if (collector is not null)
                        {
                            File.WriteAllText(Path.Combine(directory, "capture-request.json"), JsonSerializer.Serialize(new
                            {
                                executable = Path.Combine(staging, "QuickExcel.SandboxProbe.exe"), config.IPv4Port, config.IPv6Port,
                                listenerExecutable = Environment.ProcessPath, startedUtc = report["utc"]
                            }));
                            await collector.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(20));
                            report["wfpCollectorExitCode"] = collector.ExitCode;
                            complete &= collector.ExitCode == 0;
                            collector.Dispose(); collector = null;
                        }
                    }
                    if (Directory.Exists(staging)) throw new IOException("Production payload cleanup failed.");
                }
                else
                {
                    using var child = Process.Start(new ProcessStartInfo(Environment.ProcessPath!)
                    {
                        UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
                    })!;
                    using var job = Native.AttachControlJob(child);
                    try { checks = await Exchange(child.StandardInput.BaseStream, child.StandardOutput.BaseStream, child.StandardError.BaseStream, config, deadline.Token); }
                    finally { if (!child.HasExited) child.Kill(); child.WaitForExit(1000); }
                }
                deadline.Cancel();
                report[isolated ? "productionLauncher" : "positiveControl"] = new
                {
                    checks, listenerAcceptedIPv4 = await accept4, listenerAcceptedIPv6 = await accept6,
                    cleanupSucceeded = true
                };
            }
            catch (Exception e) { complete = false; report[isolated ? "productionFailure" : "controlFailure"] = e.ToString(); }
        }
        report["probeCompleted"] = complete;
        collector?.Dispose();
        // Completion is not a security verdict. Timeouts and unavailable controls remain inconclusive in JSON.
        var file = Path.Combine(directory, "network.json");
        File.WriteAllText(file, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(file);
        Console.WriteLine(File.ReadAllText(file));
        if (captureWfp && complete) return WfpEvidence.Analyze(directory) ? 0 : 2;
        return complete ? 0 : 1;
    }

    private static async Task<Dictionary<string, NetworkResult>> Exchange(Stream input, Stream output, Stream error, NetworkConfiguration config, CancellationToken token)
    {
        using var writer = new StreamWriter(input, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(output, leaveOpen: true);
        using var errors = new StreamReader(error, leaveOpen: true);
        var stderr = errors.ReadToEndAsync(token);
        await writer.WriteLineAsync(JsonSerializer.Serialize(config).AsMemory(), token);
        var line = await reader.ReadLineAsync(token);
        if (line is null) throw new IOException("Network probe exited: " + await stderr);
        _ = await stderr;
        return JsonSerializer.Deserialize<Dictionary<string, NetworkResult>>(line) ?? throw new IOException("No network results.");
    }
}
