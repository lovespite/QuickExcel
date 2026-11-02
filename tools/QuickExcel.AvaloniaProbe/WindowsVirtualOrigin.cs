using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Reflection;
using Microsoft.Web.WebView2.Core;

namespace QuickExcel.AvaloniaProbe;

/// <summary>A0-only test of the documented WebView2 virtual HTTPS API.</summary>
[SupportedOSPlatform("windows")]
internal sealed class WindowsVirtualOrigin : IDisposable
{
    private readonly object _web;
    private readonly MethodInfo _clear;
    private readonly string _host;

    internal WindowsVirtualOrigin(nint handle, string host, string directory)
    {
        // The SDK embeds its public COM interfaces with TypeIdentifier, so C#
        // cannot name them from a consuming assembly. Use their public metadata;
        // do not call the private CoreWebView2 constructor or guess vtable slots.
        var contract = typeof(CoreWebView2).Assembly.GetType("Microsoft.Web.WebView2.Core.Raw.ICoreWebView2_3", throwOnError: true)!;
        var set = contract.GetMethod("SetVirtualHostNameToFolderMapping")!;
        _clear = contract.GetMethod("ClearVirtualHostNameToFolderMapping")!;
        _web = Marshal.GetTypedObjectForIUnknown(handle, contract);
        _host = host;
        var deny = Enum.ToObject(set.GetParameters()[2].ParameterType, 0);
        try { set.Invoke(_web, [host, directory, deny]); }
        catch { Marshal.ReleaseComObject(_web); throw; }
    }

    public void Dispose()
    {
        try { _clear.Invoke(_web, [_host]); }
        finally { Marshal.ReleaseComObject(_web); }
    }
}
