using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickExcel.ScriptContracts;

// Every IPC shape is known at build time. Keep this context reflection-free so
// both the Native AOT worker and the managed host use identical wire metadata.
[JsonSourceGenerationOptions(JsonSerializerDefaults.Web,
    GenerationMode = JsonSourceGenerationMode.Metadata,
    MaxDepth = 24,
    PropertyNameCaseInsensitive = false,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    RespectNullableAnnotations = true,
    RespectRequiredConstructorParameters = true)]
[JsonSerializable(typeof(WorkerEnvelope))]
[JsonSerializable(typeof(WorkerHello))]
[JsonSerializable(typeof(LoadScript))]
[JsonSerializable(typeof(LoadedScript))]
[JsonSerializable(typeof(LoadedWorkbook))]
[JsonSerializable(typeof(BeginWorkerRound))]
[JsonSerializable(typeof(FunctionCall))]
[JsonSerializable(typeof(InvokeBatch))]
[JsonSerializable(typeof(BatchResult))]
[JsonSerializable(typeof(WorkerFailure))]
[JsonSerializable(typeof(WorkerLogs))]
[JsonSerializable(typeof(EmptyPayload))]
[JsonSerializable(typeof(LoadModules))]
[JsonSerializable(typeof(LoadedActions))]
[JsonSerializable(typeof(RunAction))]
[JsonSerializable(typeof(ActionPatch))]
[JsonSerializable(typeof(ModelCall))]
[JsonSerializable(typeof(ModelReply))]
[JsonSerializable(typeof(CompileRequest))]
[JsonSerializable(typeof(CompiledReply))]
[JsonSerializable(typeof(CompilerResult))]
[JsonSerializable(typeof(CompilerOutput[]))]
[JsonSerializable(typeof(CompilerDiagnostic[]))]
[JsonSerializable(typeof(WorkerRangeValue))]
[JsonSerializable(typeof(JsonElement))]
[JsonSerializable(typeof(object))]
[JsonSerializable(typeof(object[]))]
[JsonSerializable(typeof(List<object>))]
[JsonSerializable(typeof(Dictionary<string, object>))]
[JsonSerializable(typeof(string))]
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(byte))]
[JsonSerializable(typeof(sbyte))]
[JsonSerializable(typeof(short))]
[JsonSerializable(typeof(ushort))]
[JsonSerializable(typeof(int))]
[JsonSerializable(typeof(uint))]
[JsonSerializable(typeof(long))]
[JsonSerializable(typeof(ulong))]
[JsonSerializable(typeof(float))]
[JsonSerializable(typeof(double))]
[JsonSerializable(typeof(decimal))]
public partial class WorkerJsonContext : JsonSerializerContext;

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(CompilationDigest))]
[JsonSerializable(typeof(ModuleSetDigest))]
internal partial class DigestJsonContext : JsonSerializerContext;

/// <summary>The explicit API v2 range shape, independent of host domain types.</summary>
public sealed record WorkerRangeValue(string Kind, int Rows, int Columns, object?[] Values);

// Declaration order is part of the existing digest encoding. Never reorder
// these fields without an explicit protocol/signature compatibility decision.
internal sealed record CompilationDigest(string Compiler, string CompilerPolicy, string Options,
    string Engine, string Policy, int Api, string SourceMapMode, CompilerModule[] Modules, string[] Entries);

internal sealed record ModuleSetDigest(int ApiVersion, string[] Entries, string[] FunctionEntries,
    string[] ActionEntries, string[] EventEntries, string[] InitializationEntries, ModuleDigest[] Modules);
internal sealed record ModuleDigest(string Id, string Path, string Sha256, string Description);
