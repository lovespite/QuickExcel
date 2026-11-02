using System.Security.Cryptography;
using System.Text;
using QuickExcel.Security;

namespace QuickExcel.Core.Test;

public sealed class PlatformTrustProtectionTests
{
    private sealed class TestVault : ISecretKeyVault
    {
        private readonly byte[] _key=RandomNumberGenerator.GetBytes(32);
        public byte[] Get(string identity,bool create)=>_key.ToArray();
    }
    private sealed class UnavailableVault : ISecretKeyVault
    {public byte[] Get(string identity,bool create)=>throw new UserProtectionUnavailableException("Locked test key service");}
    [Fact]
    public void VaultCipherAuthenticatesEveryByteAndTheExactPurpose()
    {
        var protection=new VaultProtection(new TestVault(),"test-only");
        var data=Encoding.UTF8.GetBytes("controlled trusted certificate records");
        var purpose=Encoding.UTF8.GetBytes("/isolated/Store.json");
        var encrypted=protection.Protect(data,purpose);
        Assert.Equal(data,protection.Unprotect(encrypted,purpose));
        Assert.False(encrypted.AsSpan().IndexOf(data)>=0);
        Assert.NotEqual(encrypted,protection.Protect(data,purpose));
        foreach(var index in new[]{0,1,12,13,28,29,encrypted.Length-1})
        {
            var changed=encrypted.ToArray();changed[index]^=1;
            Assert.ThrowsAny<CryptographicException>(()=>protection.Unprotect(changed,purpose));
        }
        Assert.ThrowsAny<CryptographicException>(()=>protection.Unprotect(encrypted,Encoding.UTF8.GetBytes("/isolated/store.json")));
        Assert.ThrowsAny<CryptographicException>(()=>new VaultProtection(new TestVault(),"test-only").Unprotect(encrypted,purpose));
    }
    [Fact]
    public void MissingKeyServiceNeverCreatesAnUnprotectedFallback()
    {
        var protection=new VaultProtection(new UnavailableVault(),"test-only");
        Assert.Throws<UserProtectionUnavailableException>(()=>protection.Protect([1,2,3],[4]));
        Assert.Throws<UserProtectionUnavailableException>(()=>protection.Unprotect(new byte[29]{1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0},[4]));
    }
}
