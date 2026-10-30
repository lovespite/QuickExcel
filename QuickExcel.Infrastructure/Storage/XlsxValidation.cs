using System.Collections.Immutable;
using System.Xml.Linq;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static class XlsxValidation
{
    private static readonly string[] Types = ["none", "whole", "decimal", "list", "date", "textLength", "custom"];
    private static readonly string[] Operators = ["between", "notBetween", "equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"];
    private static readonly string[] Alerts = ["stop", "warning", "information"];
    public static ImmutableArray<WorksheetValidation> Read(XlsxPackage package, string path, XElement sheet)
    {
        var result = ImmutableArray.CreateBuilder<WorksheetValidation>();
        foreach (var extension in sheet.Descendants().Where(e => e.Name.LocalName == "dataValidations" && e.Name.Namespace != Main))
            package.Loss("扩展数据验证尚未建模，原规则保留并限制编辑。", "worksheet/dataValidation", path, extension);
        if (sheet.Elements(Main + "dataValidations").Count() > 1) throw new InvalidDataException("数据验证集合重复。");
        var root = sheet.Element(Main + "dataValidations"); if (root is null) return [];
        try
        {
            XlsxTables.Shape(root, "dataValidation", "count disablePrompts xWindow yWindow");
            // Prompt placement is application UI state. disablePrompts changes visible behavior.
            if (Bool((string?)root.Attribute("disablePrompts"))) throw new NotSupportedException("整表禁用验证提示尚未建模。");
            foreach (var e in root.Elements())
            {
                XlsxTables.Shape(e, "formula1 formula2", "type errorStyle operator allowBlank showDropDown showInputMessage showErrorMessage errorTitle error promptTitle prompt sqref imeMode");
                if ((string?)e.Attribute("imeMode") is { } ime && ime != "noControl") throw new NotSupportedException("验证 IME 模式尚未支持。");
                var kind = Array.IndexOf(Types, (string?)e.Attribute("type") ?? "none");
                var op = Array.IndexOf(Operators, (string?)e.Attribute("operator") ?? "between");
                var alert = Array.IndexOf(Alerts, (string?)e.Attribute("errorStyle") ?? "stop");
                if (kind < 0 || op < 0 || alert < 0) throw new NotSupportedException("数据验证类型、比较符或错误样式未支持。");
                string? Formula(string name)
                {
                    if (e.Elements(Main + name).Count() > 1) throw new InvalidDataException("数据验证公式重复。");
                    var f = e.Element(Main + name); if (f is null) return null;
                    XlsxTables.Shape(f, "", ""); return XlsxStructuredFormula.Read("=" + Decode(f.Value).TrimStart('='));
                }
                string? Text(string name) => e.Attribute(name) is { } a ? Decode(a.Value) : null;
                var ranges = Attr(e, "sqref").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse);
                result.Add(new(ranges, (ValidationKind)kind, Formula("formula1"), Formula("formula2"), (ValidationOperator)op,
                    Bool((string?)e.Attribute("allowBlank")), !Bool((string?)e.Attribute("showDropDown")),
                    Bool((string?)e.Attribute("showInputMessage")), Bool((string?)e.Attribute("showErrorMessage")), (ValidationAlert)alert,
                    Text("promptTitle"), Text("prompt"), Text("errorTitle"), Text("error")));
                if (result.Count > 1024) throw new InvalidDataException("数据验证规则超过 1024。");
            }
            // Validate overlap and aggregate budgets before accepting any rule.
            _ = new WorksheetSnapshot(SheetId.New(), "Validation", data: new() { Validations = result.ToImmutable() });
            return result.ToImmutable();
        }
        catch (Exception ex) when (ex is NotSupportedException or FormatException or ArgumentException)
        {
            package.Loss("数据验证无法安全执行，保留原规则并限制编辑：" + ex.Message, "worksheet/dataValidation", path, root);
            return [];
        }
    }
    public static XElement? Write(WorksheetSnapshot sheet)
    {
        if (sheet.Data.Validations.IsEmpty) return null;
        return new(Main + "dataValidations", new XAttribute("count", sheet.Data.Validations.Length), sheet.Data.Validations.Select(v =>
            new XElement(Main + "dataValidation", new XAttribute("type", Types[(int)v.Kind]), new XAttribute("operator", Operators[(int)v.Comparison]),
                new XAttribute("errorStyle", Alerts[(int)v.Alert]), new XAttribute("allowBlank", v.AllowBlank ? 1 : 0),
                new XAttribute("showDropDown", v.ShowDropDown ? 0 : 1), new XAttribute("showInputMessage", v.ShowInput ? 1 : 0),
                new XAttribute("showErrorMessage", v.ShowError ? 1 : 0), new XAttribute("sqref", string.Join(" ", v.Ranges)),
                v.InputTitle is null ? null : new XAttribute("promptTitle", Encode(v.InputTitle)), v.InputMessage is null ? null : new XAttribute("prompt", Encode(v.InputMessage)),
                v.ErrorTitle is null ? null : new XAttribute("errorTitle", Encode(v.ErrorTitle)), v.ErrorMessage is null ? null : new XAttribute("error", Encode(v.ErrorMessage)),
                v.Formula1 is null ? null : new XElement(Main + "formula1", Encode(XlsxStructuredFormula.Write(v.Formula1).TrimStart('='))),
                v.Formula2 is null ? null : new XElement(Main + "formula2", Encode(XlsxStructuredFormula.Write(v.Formula2).TrimStart('='))))));
    }
}
