using System.Text.Json.Serialization;

namespace QuickExcel.Scripting;

// The native bridge has a closed set of message shapes. Source-generated metadata
// keeps those shapes available when the worker is trimmed and compiled ahead of time.
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true)]
[JsonSerializable(typeof(string))]
[JsonSerializable(typeof(string[]))]
[JsonSerializable(typeof(JsFunctionDescriptor[]))]
[JsonSerializable(typeof(JsActionDescriptor[]))]
[JsonSerializable(typeof(PrepareModulesOptions))]
[JsonSerializable(typeof(NativeModuleBundle))]
[JsonSerializable(typeof(HostCallFailure))]
internal sealed partial class ScriptingJsonContext : JsonSerializerContext;

internal sealed record PrepareModulesOptions(bool Share, bool Automatic);
internal sealed record NativeModule(string Id, string Source, int Roles);
internal sealed record NativeModuleBundle(NativeModule[] Modules, string[] Entries);
internal sealed record HostCallFailure(bool Ok, string Error, string Code);
