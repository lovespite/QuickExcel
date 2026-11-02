using System.Text.Json;
using System.Text.Json.Nodes;
using QuickExcel.Domain;
using QuickExcel.Security;

internal static class TrustChecks
{
    internal static int Run(string directory,bool unavailable)
    {
        if(Directory.Exists(directory))return 2;
        Directory.CreateDirectory(directory);directory=Path.GetFullPath(directory);
        var checks=new List<string>();Exception? failure=null;
        var previous=Environment.GetEnvironmentVariable("DBUS_SESSION_BUS_ADDRESS");
        try
        {
            using var certificate=ScriptSignatures.CreateSelfSigned("Controlled platform trust check");
            var modules=new ScriptModuleSet(new Dictionary<string,ScriptModule>{["main"]=new("scripts/main.js","quickExcel.functions.register('TEST',()=>42);")},["main"],scriptApiVersion:2);
            modules=modules.WithSignature(ScriptSignatures.Sign(modules,certificate));
            var path=Path.Combine(directory,"trust.json");var store=new ScriptCertificateTrustStore(path);
            if(store.Verify(modules).CanAutoEnable)throw new Exception("Unexpected initial auto-trust.");
            if(unavailable)
            {
                if(!OperatingSystem.IsLinux())return 2;
                Environment.SetEnvironmentVariable("DBUS_SESSION_BUS_ADDRESS","unix:path="+Path.Combine(directory,"absent-bus"));
                var refused=false;try{store.Trust(modules);}catch(UserProtectionUnavailableException){refused=true;}
                if(!refused || File.Exists(path))throw new Exception("Unavailable key service wrote trust data.");
                if(store.Verify(modules).CanAutoEnable)throw new Exception("Unavailable service granted auto-trust.");
                checks.Add("Missing Secret Service refuses persistent trust without plaintext fallback or automatic authorization");
            }
            else
            {
                store.Trust(modules);
                var text=File.ReadAllText(path);var original=File.ReadAllBytes(path);
                var envelope=JsonNode.Parse(text)!.AsObject();
                if(envelope["Version"]!.GetValue<int>()!=2 || text.Contains(certificate.Subject) || text.Contains(store.Read().Single().Fingerprint))throw new Exception("Invalid protected Unix envelope.");
                if(!new ScriptCertificateTrustStore(path).Verify(modules).CanAutoEnable)throw new Exception("Persisted pin not recovered.");
                checks.Add("Signed certificate pin round-trip through OS key service and authenticated v2 storage");
                File.Copy(path,Path.Combine(directory,"copied.json"));
                var rejected=false;try{new ScriptCertificateTrustStore(Path.Combine(directory,"copied.json")).Read();}catch(Exception ex)when(ex is UserProtectionUnavailableException or InvalidDataException){rejected=true;}
                if(!rejected)throw new Exception("Trust data was portable to another store.");
                checks.Add("Store path binding rejects copied ciphertext");
                var data=Convert.FromBase64String(envelope["Data"]!.GetValue<string>());data[^1]^=1;envelope["Data"]=Convert.ToBase64String(data);File.WriteAllText(path,envelope.ToJsonString());
                rejected=false;try{store.Read();}catch(InvalidDataException){rejected=true;}if(!rejected)throw new Exception("Modified ciphertext accepted.");
                checks.Add("Tampered ciphertext rejected");
                File.WriteAllBytes(path,original);store.Remove(store.Read().Single().Fingerprint);
                if(store.Verify(modules).CanAutoEnable || store.Read().Count!=0)throw new Exception("Revoked pin remained trusted.");
                checks.Add("Revocation persists and removes automatic authorization");
            }
        }
        catch(Exception ex){failure=ex;}
        finally{Environment.SetEnvironmentVariable("DBUS_SESSION_BUS_ADDRESS",previous);}
        File.WriteAllText(Path.Combine(directory,"result.json"),JsonSerializer.Serialize(new{passed=failure is null,checks,error=failure?.ToString()},new JsonSerializerOptions{WriteIndented=true}));
        Console.WriteLine(failure?.ToString() ?? "PASS: platform trust checks");return failure is null ? 0 : 1;
    }
}
