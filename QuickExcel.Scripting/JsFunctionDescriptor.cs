namespace QuickExcel.Scripting;

public sealed record JsFunctionDescriptor(string Name, int ApiVersion, int MinArguments, int MaxArguments,
    string ReturnType, string Profile);