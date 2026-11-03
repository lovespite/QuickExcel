using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace QuickExcel.Scripting;

/// <summary>Called only by the isolated compilation role; never executes workbook source.</summary>
public static class JscCompilerSession
{
    public static string Compile(string requestJson, string expectedIdentity)
    {
        if (Encoding.UTF8.GetByteCount(requestJson) > 16 * 1024 * 1024)
            throw new InvalidOperationException("Compiler bridge input exceeds its budget.");
        var session = CreateVerifiedContext(expectedIdentity);
        try
        {
            // Release the large transient bundle/JSON/UTF-16 bootstrap copies
            // before native Program construction. CLR heap pressure cannot see
            // JSC allocations in the same fixed 1 GiB compiler process budget.
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
            var result = Call(session, "compile", "[" + requestJson + "]", out var error);
            if (result == IntPtr.Zero) throw ReadError(error);
            try { return Marshal.PtrToStringUTF8(result) ?? throw new InvalidOperationException("Missing compiler result."); }
            finally { Free(result); }
        }
        finally { Destroy(session); }
    }

    private static IntPtr CreateVerifiedContext(string expectedIdentity)
    {
        var file = Path.Combine(AppContext.BaseDirectory, "typescript.bundle.json");
        var info = new FileInfo(file);
        if (info.Length > 16 * 1024 * 1024) throw new InvalidOperationException("Compiler payload exceeds its budget.");
        var data = File.ReadAllBytes(file);
        if (Convert.ToHexString(SHA256.HashData(data)) != expectedIdentity || Marshal.PtrToStringUTF8(Engine()) != JscSession.EngineIdentity)
            throw new InvalidOperationException("Compiler payload or native build differs from the accepted identity.");
        using var payload = JsonDocument.Parse(data);
        var p = payload.RootElement;
        if (p.GetProperty("version").GetString() != "5.9.3") throw new InvalidOperationException("Unsupported compiler.");
        var bootstrap = p.GetProperty("compiler").GetString() + "\n" + p.GetProperty("frontend").GetString() +
            "\n({compile(request) {return qxCompile({...request, mode:'check', libraries:" + p.GetProperty("libraries").GetRawText() +
            ", declarations:" + p.GetProperty("declarations").GetRawText() + "});}})";
        if (Encoding.UTF8.GetByteCount(bootstrap) > 16 * 1024 * 1024)
            throw new InvalidOperationException("Compiler bridge input exceeds its budget.");
        var session = Create(bootstrap, out var error);
        if (session == IntPtr.Zero) throw ReadError(error);
        return session;
    }
    private static InvalidOperationException ReadError(IntPtr error)
    {
        try { return new(Marshal.PtrToStringUTF8(error) ?? "Compilation failed."); }
        finally { if (error != IntPtr.Zero) Free(error); }
    }
    [DllImport("JavaScriptCore", EntryPoint = "qx_create_compiler", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr Create([MarshalAs(UnmanagedType.LPUTF8Str)] string bootstrap, out IntPtr error);
    [DllImport("JavaScriptCore", EntryPoint = "qx_call", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr Call(IntPtr session, [MarshalAs(UnmanagedType.LPUTF8Str)] string method, [MarshalAs(UnmanagedType.LPUTF8Str)] string json, out IntPtr error);
    [DllImport("JavaScriptCore", EntryPoint = "qx_destroy", CallingConvention = CallingConvention.Cdecl)]
    private static extern void Destroy(IntPtr session);
    [DllImport("JavaScriptCore", EntryPoint = "qx_free_string", CallingConvention = CallingConvention.Cdecl)]
    private static extern void Free(IntPtr text);
    [DllImport("JavaScriptCore", EntryPoint = "qx_engine_identity", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr Engine();
}
