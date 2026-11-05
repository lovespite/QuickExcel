using QuickExcel.Core;

namespace QuickExcel.Runtime;

/// <summary>Process-wide host policy, independent of workbook signatures and trust.</summary>
public sealed class ScriptExecutionPolicy
{
    public const string EnvironmentVariable = "QUICKEXCEL_NO_SCRIPTS";
    public static ScriptExecutionPolicy Current { get; } = Read(Environment.GetCommandLineArgs().Skip(1),
        target => target == EnvironmentVariableTarget.Process || OperatingSystem.IsWindows()
            ? Environment.GetEnvironmentVariable(EnvironmentVariable, target) : null);

    private readonly object _gate = new();
    private readonly bool _commandLine, _environment;
    private bool _userSetting;
    private CancellationTokenSource _lifetime = new();

    private ScriptExecutionPolicy(bool commandLine, bool environment)
    { _commandLine = commandLine; _environment = environment; }

    // Read each scope independently: a process/user value of 0 must not override a machine value of 1.
    public static ScriptExecutionPolicy Read(IEnumerable<string> arguments, Func<EnvironmentVariableTarget, string?> readEnvironment) =>
        new(arguments.Contains("--no-scripts", StringComparer.Ordinal),
            new[] { EnvironmentVariableTarget.Process, EnvironmentVariableTarget.User, EnvironmentVariableTarget.Machine }
                .Select(readEnvironment).Any(value => value == "1"));

    public bool UserSetting
    {
        get { lock (_gate) return _userSetting; }
        set
        {
            CancellationTokenSource? revoked = null;
            lock (_gate)
            {
                if (_userSetting == value) return;
                _userSetting = value;
                if (value) revoked = _lifetime;
                else if (!_commandLine && !_environment) _lifetime = new();
            }
            // Cancel outside the lock: worker cleanup may call back into the policy.
            revoked?.Cancel();
        }
    }

    public bool IsDisabled => UserSetting || _commandLine || _environment;
    public string Status => !IsDisabled ? "" : "No scripts - " + string.Join("/",
        new[] { UserSetting ? "User setting" : null, _commandLine ? "Command line" : null, _environment ? "Env" : null }.OfType<string>());

    public void RequireEnabled()
    {
        if (IsDisabled) throw new InvalidOperationException(MessageText.Get("当前环境已禁用脚本。") + " " + Status);
    }

    internal CancellationToken CaptureLifetime()
    {
        lock (_gate) { RequireEnabled(); return _lifetime.Token; }
    }
}
