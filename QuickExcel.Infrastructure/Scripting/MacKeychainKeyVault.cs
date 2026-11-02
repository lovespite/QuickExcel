using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace QuickExcel.Security;

/// <summary>Data Protection Keychain, local device only; unavailable/locked means no persistent trust.</summary>
internal sealed class MacKeychainKeyVault : ISecretKeyVault
{
    private const string Security="/System/Library/Frameworks/Security.framework/Security";
    private const string CoreFoundation="/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
    public byte[] Get(string identity,bool create)
    {
        if(!OperatingSystem.IsMacOS())throw new PlatformNotSupportedException();
        using var query=new Query(identity);
        query.Set("kSecReturnData",query.True);
        var status=SecItemCopyMatching(query.Dictionary,out var result);
        if(status==0)return Read(result);
        if(status!=-25300 || !create)throw Unavailable(status);
        var key=RandomNumberGenerator.GetBytes(32);
        try
        {
            using var insert=new Query(identity);
            insert.Set("kSecAttrAccessible",insert.Constant("kSecAttrAccessibleWhenUnlockedThisDeviceOnly"));
            var data=CFDataCreate(0,key,key.Length); if(data==0)throw new OutOfMemoryException();
            try {insert.Set("kSecValueData",data);status=SecItemAdd(insert.Dictionary,out result); if(result!=0)CFRelease(result);}
            finally {CFRelease(data);}
            if(status==0)return key;
            if(status==-25299) {CryptographicOperations.ZeroMemory(key);return Get(identity,false);}
            throw Unavailable(status);
        }
        catch {CryptographicOperations.ZeroMemory(key);throw;}
    }
    private static byte[] Read(nint data)
    {
        try
        {
            if(data==0 || CFGetTypeID(data)!=CFDataGetTypeID() || CFDataGetLength(data)!=32)throw Unavailable(-1);
            var key=new byte[32]; Marshal.Copy(CFDataGetBytePtr(data),key,0,key.Length);return key;
        }
        finally {if(data!=0)CFRelease(data);}
    }
    private static UserProtectionUnavailableException Unavailable(int status)=>new($"macOS Keychain 不可用、已锁定或应用签名权限未就绪（{status}）；持久信任关闭，仍可明确选择仅本次允许。");
    private sealed class Query : IDisposable
    {
        private nint _security,_core;
        internal nint Dictionary {get;}
        internal nint True=>Marshal.ReadIntPtr(NativeLibrary.GetExport(_core,"kCFBooleanTrue"));
        internal nint Constant(string name)=>Marshal.ReadIntPtr(NativeLibrary.GetExport(_security,name));
        internal void Set(string key,nint value)=>CFDictionarySetValue(Dictionary,Constant(key),value);
        internal Query(string identity)
        {
            try
            {
            _security=NativeLibrary.Load(Security);_core=NativeLibrary.Load(CoreFoundation);
            Dictionary=CFDictionaryCreateMutable(0,0,NativeLibrary.GetExport(_core,"kCFTypeDictionaryKeyCallBacks"),NativeLibrary.GetExport(_core,"kCFTypeDictionaryValueCallBacks"));
            if(Dictionary==0)throw new OutOfMemoryException();
            Set("kSecClass",Constant("kSecClassGenericPassword"));
            Set("kSecUseDataProtectionKeychain",True);
            Set("kSecAttrSynchronizable",Marshal.ReadIntPtr(NativeLibrary.GetExport(_core,"kCFBooleanFalse")));
            Set("kSecUseAuthenticationUI",Constant("kSecUseAuthenticationUIFail"));
            AddText("kSecAttrService","QuickExcel.ScriptTrust.v2"); AddText("kSecAttrAccount",identity);
            }
            catch {Dispose();throw;}
        }
        private void AddText(string key,string value)
        {var text=CFStringCreateWithCString(0,value,0x08000100);if(text==0)throw new OutOfMemoryException();try{Set(key,text);}finally{CFRelease(text);}}
        public void Dispose(){if(Dictionary!=0)CFRelease(Dictionary);if(_security!=0)NativeLibrary.Free(_security);if(_core!=0)NativeLibrary.Free(_core);}
    }
    [DllImport(Security)] private static extern int SecItemCopyMatching(nint query,out nint result);
    [DllImport(Security)] private static extern int SecItemAdd(nint attributes,out nint result);
    [DllImport(CoreFoundation)] private static extern nint CFDictionaryCreateMutable(nint allocator,nint capacity,nint keys,nint values);
    [DllImport(CoreFoundation)] private static extern void CFDictionarySetValue(nint dictionary,nint key,nint value);
    [DllImport(CoreFoundation)] private static extern nint CFStringCreateWithCString(nint allocator,[MarshalAs(UnmanagedType.LPUTF8Str)]string text,uint encoding);
    [DllImport(CoreFoundation)] private static extern nint CFDataCreate(nint allocator,byte[] bytes,nint count);
    [DllImport(CoreFoundation)] private static extern nint CFDataGetLength(nint data);
    [DllImport(CoreFoundation)] private static extern nint CFDataGetBytePtr(nint data);
    [DllImport(CoreFoundation)] private static extern nuint CFGetTypeID(nint item);
    [DllImport(CoreFoundation)] private static extern nuint CFDataGetTypeID();
    [DllImport(CoreFoundation)] private static extern void CFRelease(nint item);
}
