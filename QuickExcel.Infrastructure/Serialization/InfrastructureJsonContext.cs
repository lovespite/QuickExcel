using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;
using QuickExcel.Security;
using QuickExcel.Storage;

namespace QuickExcel.Serialization;

// Each persisted format keeps its original options. In particular, cache/signature
// metadata uses the strict IPC camel-case format; preferences retain PascalCase.
[JsonSourceGenerationOptions(GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(Dictionary<string, string>))]
[JsonSerializable(typeof(string))]
[JsonSerializable(typeof(string[]))]
[JsonSerializable(typeof(JsonObject))]
[JsonSerializable(typeof(JsonNode))]
[JsonSerializable(typeof(ScriptSignature))]
[JsonSerializable(typeof(ScriptCertificateTrustStore.ExecutionPolicyIdentity))]
[JsonSerializable(typeof(ScriptCompilationCache.Authentication))]
[JsonSerializable(typeof(ScriptSigningPreferences))]
[JsonSerializable(typeof(ScriptSigningPreferencesStore.Envelope), TypeInfoPropertyName = "SigningEnvelope")]
[JsonSerializable(typeof(ScriptCertificateTrustStore.ProtectedTrust))]
[JsonSerializable(typeof(ScriptCertificateTrustStore.TrustState))]
[JsonSerializable(typeof(TrustedScriptCertificate[]))]
[JsonSerializable(typeof(ScriptCompilationCache.Envelope), TypeInfoPropertyName = "CompilationEnvelope")]
[JsonSerializable(typeof(WorkbookRecoveryStore.Owner))]
[JsonSerializable(typeof(RecoveryMetadata))]
[JsonSerializable(typeof(LinuxWorkerLease.Record), TypeInfoPropertyName = "LinuxLeaseRecord")]
[JsonSerializable(typeof(MacWorkerLease.Record), TypeInfoPropertyName = "MacLeaseRecord")]
internal partial class InfrastructureJsonContext : JsonSerializerContext
{
    internal static InfrastructureJsonContext Strict { get; } = new(new JsonSerializerOptions
    { UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow });
    internal static InfrastructureJsonContext Protocol { get; } = new(new JsonSerializerOptions(WorkerProtocol.Json));
}
