using System.Text.RegularExpressions;
namespace QuickExcel.Core;

public sealed record ScriptFunctionDescriptor(string Name, int ApiVersion = 1, int MinArguments = 0,
    int MaxArguments = 255, string ReturnType = "any", string Profile = "conservative")
{
    public void Validate()
    {
        if (ApiVersion is not (1 or 2) || !Regex.IsMatch(Name, @"^[A-Z_][A-Z0-9_]*(\.[A-Z_][A-Z0-9_]*)*$")
            || ApiVersion == 2 && !Name.StartsWith("JS.", StringComparison.Ordinal) ||
            MinArguments < 0 || MaxArguments < MinArguments || MaxArguments > 255 ||
            ReturnType is not ("any" or "number" or "text" or "boolean" or "blank") || Profile != "conservative")
            throw new ArgumentException("函数描述符无效；M3 仅支持 conservative 执行策略。");
    }
}
