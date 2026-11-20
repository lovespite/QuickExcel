# WinVerifyTrust distinguishes an untrusted self-signed root from a corrupt PE
# signature without importing certificates into the user's trust stores.
if (-not ('QuickExcel.Publishing.Authenticode' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
namespace QuickExcel.Publishing {
    // ClickOnce uses the original xmldsig namespace for its SHA-256 URIs.
    // Modern SignedXml only registers the W3C xmldsig-more aliases by default.
    public sealed class ClickOnceSha256SignatureDescription : SignatureDescription {
        public ClickOnceSha256SignatureDescription() { KeyAlgorithm = typeof(RSA).AssemblyQualifiedName; }
        public override HashAlgorithm CreateDigest() => SHA256.Create();
        public override AsymmetricSignatureFormatter CreateFormatter(AsymmetricAlgorithm key) {
            var formatter = new RSAPKCS1SignatureFormatter(key);
            formatter.SetHashAlgorithm("SHA256"); return formatter;
        }
        public override AsymmetricSignatureDeformatter CreateDeformatter(AsymmetricAlgorithm key) {
            var formatter = new RSAPKCS1SignatureDeformatter(key);
            formatter.SetHashAlgorithm("SHA256"); return formatter;
        }
    }
    public static class Authenticode {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct FileInfo {
            public uint Size;
            [MarshalAs(UnmanagedType.LPWStr)] public string Path;
            public IntPtr Handle, KnownSubject;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct TrustData {
            public uint Size;
            public IntPtr Policy, Sip;
            public uint UIChoice, RevocationChecks, UnionChoice;
            public IntPtr File;
            public uint StateAction;
            public IntPtr State, Url;
            public uint ProviderFlags, UIContext;
            public IntPtr SignatureSettings;
        }
        [DllImport("wintrust.dll", ExactSpelling = true, PreserveSig = true)]
        static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);
        public static int Verify(string path) {
            var file = new FileInfo { Size = (uint)Marshal.SizeOf<FileInfo>(), Path = path };
            var pointer = Marshal.AllocHGlobal(Marshal.SizeOf<FileInfo>());
            Marshal.StructureToPtr(file, pointer, false);
            var data = new TrustData { Size = (uint)Marshal.SizeOf<TrustData>(), UIChoice = 2,
                UnionChoice = 1, File = pointer, StateAction = 1, ProviderFlags = 0x1010 };
            var action = new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");
            try { return WinVerifyTrust(new IntPtr(-1), ref action, ref data); }
            finally {
                data.StateAction = 2;
                WinVerifyTrust(new IntPtr(-1), ref action, ref data);
                Marshal.DestroyStructure<FileInfo>(pointer);
                Marshal.FreeHGlobal(pointer);
            }
        }
    }
}
'@
}
function Assert-QuickExcelCodeSignature([string]$Path, [string]$Thumbprint) {
    $resolved = (Get-Item -LiteralPath $Path).FullName
    $signature = Get-AuthenticodeSignature -LiteralPath $resolved
    $result = [QuickExcel.Publishing.Authenticode]::Verify($resolved)
    # CERT_E_UNTRUSTEDROOT is expected for a self-signed certificate. Every other
    # error, including TRUST_E_BAD_DIGEST, wrong usage and expiry, remains fatal.
    if (-not $signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne $Thumbprint -or $result -notin @(0, -2146762487)) {
        throw ('Invalid Authenticode signature: {0}; HRESULT 0x{1:X8}' -f $resolved,$result)
    }
}
function Assert-QuickExcelWebView2Bootstrapper([string]$Path) {
    $file = Get-Item -LiteralPath $Path
    $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation(?:,|$)' -or $file.Length -lt 100KB -or $file.Length -gt 10MB) {
        throw 'A small, Microsoft-signed WebView2 Evergreen online bootstrapper is required.'
    }
}
