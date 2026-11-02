using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace QuickExcel.Security;

public sealed class UserProtectionUnavailableException(string message, Exception? inner = null) : IOException(message, inner);

internal abstract class TrustDataProtection
{
    internal abstract string Id { get; }
    internal abstract byte[] Protect(byte[] data, byte[] purpose);
    internal abstract byte[] Unprotect(byte[] data, byte[] purpose);
    internal static TrustDataProtection Create() => OperatingSystem.IsWindows() ? new WindowsProtection()
        : OperatingSystem.IsLinux() ? new VaultProtection(new SecretServiceKeyVault(), "secret-service-aesgcm-v1")
        : OperatingSystem.IsMacOS() ? new VaultProtection(new MacKeychainKeyVault(), "keychain-aesgcm-v1")
        : throw new PlatformNotSupportedException("当前平台不提供持久脚本信任保护。");

    private sealed class WindowsProtection : TrustDataProtection
    {
        internal override string Id => "dpapi-current-user-v1";
        internal override byte[] Protect(byte[] data, byte[] purpose) => OperatingSystem.IsWindows()
            ? ProtectedData.Protect(data,purpose,DataProtectionScope.CurrentUser) : throw new PlatformNotSupportedException();
        internal override byte[] Unprotect(byte[] data, byte[] purpose) => OperatingSystem.IsWindows()
            ? ProtectedData.Unprotect(data,purpose,DataProtectionScope.CurrentUser) : throw new PlatformNotSupportedException();
    }
}

internal interface ISecretKeyVault { byte[] Get(string identity, bool create); }

internal sealed class VaultProtection(ISecretKeyVault vault, string id) : TrustDataProtection
{
    internal override string Id => id;
    internal override byte[] Protect(byte[] data, byte[] purpose)
    {
        var key=vault.Get(Convert.ToHexString(SHA256.HashData(purpose)),true);
        try
        {
            var result=new byte[29+data.Length]; result[0]=1;
            RandomNumberGenerator.Fill(result.AsSpan(1,12));
            using var cipher=new AesGcm(key,16);
            cipher.Encrypt(result.AsSpan(1,12),data,result.AsSpan(29),result.AsSpan(13,16),purpose);
            return result;
        }
        finally {CryptographicOperations.ZeroMemory(key);}
    }
    internal override byte[] Unprotect(byte[] data, byte[] purpose)
    {
        if(data.Length<29 || data[0]!=1)throw new CryptographicException("Invalid protected trust envelope.");
        var key=vault.Get(Convert.ToHexString(SHA256.HashData(purpose)),false);
        var result=new byte[data.Length-29];
        try
        {
            using var cipher=new AesGcm(key,16);
            cipher.Decrypt(data.AsSpan(1,12),data.AsSpan(29),data.AsSpan(13,16),result,purpose);
            return result;
        }
        catch {CryptographicOperations.ZeroMemory(result);throw;}
        finally {CryptographicOperations.ZeroMemory(key);}
    }
}

/// <summary>Keys remain in the desktop Secret Service. No key file or plaintext fallback.</summary>
internal sealed class SecretServiceKeyVault : ISecretKeyVault
{
    public byte[] Get(string identity,bool create)
    {
        if(!OperatingSystem.IsLinux() || !File.Exists("/usr/bin/secret-tool"))throw Unavailable();
        var existing=Execute(["lookup","application","QuickExcel.ScriptTrust.v2","store",identity],null);
        if(existing.Length!=0)return Decode(existing);
        if(!create)throw Unavailable();
        var key=RandomNumberGenerator.GetBytes(32);
        try
        {
            Execute(["store","--label=QuickExcel script certificate trust key","application","QuickExcel.ScriptTrust.v2","store",identity],Convert.ToBase64String(key));
            var saved=Decode(Execute(["lookup","application","QuickExcel.ScriptTrust.v2","store",identity],null));
            try {if(!CryptographicOperations.FixedTimeEquals(key,saved))throw Unavailable();}
            finally {CryptographicOperations.ZeroMemory(saved);}
            return key;
        }
        catch {CryptographicOperations.ZeroMemory(key);throw;}
    }
    private static byte[] Decode(string value)
    {
        try {var bytes=Convert.FromBase64String(value); if(bytes.Length==32)return bytes; CryptographicOperations.ZeroMemory(bytes);}
        catch(FormatException) { }
        throw new UserProtectionUnavailableException("系统密钥服务中的 QuickExcel 密钥无效；持久信任保持关闭。");
    }
    private static string Execute(string[] arguments,string? input)
    {
        var start=new ProcessStartInfo("/usr/bin/secret-tool") {UseShellExecute=false,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true};
        foreach(var argument in arguments)start.ArgumentList.Add(argument);
        using var process=Process.Start(start) ?? throw Unavailable();
        using var timeout=new CancellationTokenSource(TimeSpan.FromSeconds(input is null ? 5 : 15));
        try
        {
            var output=process.StandardOutput.ReadToEndAsync(timeout.Token); var error=process.StandardError.ReadToEndAsync(timeout.Token);
            if(input is not null)process.StandardInput.WriteLine(input);
            process.StandardInput.Close();
            process.WaitForExitAsync(timeout.Token).GetAwaiter().GetResult();
            var text=output.GetAwaiter().GetResult().Trim(); var diagnostic=error.GetAwaiter().GetResult();
            if(text.Length>128 || diagnostic.Length>0 || (process.ExitCode!=0 && !(input is null && process.ExitCode==1 && text.Length==0)))throw Unavailable();
            return text;
        }
        catch(OperationCanceledException ex){throw new UserProtectionUnavailableException("系统密钥服务未响应；本次未保存持久信任，可解锁密钥环后重试。",ex);}
        finally {if(!process.HasExited){process.Kill();process.WaitForExit(1000);}}
    }
    private static UserProtectionUnavailableException Unavailable()=>new("系统密钥服务不可用或已锁定；持久自动信任关闭，仍可明确选择仅本次允许。Linux 需要 libsecret-tools 和已解锁的 Secret Service。");
}
