using QuickExcel.Domain;
using System.Security.Cryptography;
using System.Text;

namespace QuickExcel.Presentation.Web;

/// <summary>Editable starter source; registration happens only when the user applies scripts.</summary>
public static class ScriptModuleTemplates
{
    public static string CreateConventional(string id, ScriptLanguage language, string? role)
    {
        var typed = language == ScriptLanguage.TypeScript;
        if (role == "events") return "// Use Add Event Handler to insert a synchronous exported handler.\n// Handlers are discovered automatically when scripts are applied.\n";
        if (role == "initialization") return "export function main(" + (typed ? "workbook: QuickExcel.Workbook, host: QuickExcel.Host" : "workbook, host") + ") {\n    // Initialize shared Actions / Events state here.\n}\n";
        if (role == "formula-functions")
        {
            var stem = CompactName(string.Concat(id.Select(c => char.IsAsciiLetterLower(c) || char.IsAsciiDigit(c) ? c.ToString() : $"_{(int)c:X2}")));
            if (char.IsAsciiDigit(stem[0])) stem = "_" + stem;
            return $$"""
                // Apply scripts, then try =JS.{{stem}}.DOUBLE(21) in a cell.
                // Exported functions are discovered automatically. UDFs use only their arguments.
                {{(typed ? "" : "/** @param {unknown} value @returns {number} */\n")}}export function DOUBLE(value{{(typed ? ": unknown" : "")}}){{(typed ? ": number" : "")}} {
                    if (typeof value !== 'number' || !Number.isFinite(value)) throw new TypeError('Expected a finite number.');
                    return value * 2;
                }

                """;
        }
        return role == "actions"
            ? "// One exported main defines this action; writes form one undoable transaction.\nexport function main(" +
                (typed ? "workbook: QuickExcel.Workbook, host: QuickExcel.Host" : "workbook, host") +
                ") {\n    workbook.getActiveCell().setValue(42);\n}\n"
            : "// Import this Common module by its logical ID: '" + id + "'.\nexport function example(value" +
                (typed ? ": number" : "") + ") {\n    return value;\n}\n";
    }

    public static string Create(string id, ScriptLanguage language, bool functions, bool actions)
    {
        var typed = language == ScriptLanguage.TypeScript;
        var sections = new List<string>();
        if (functions)
        {
            // UDF names are case-insensitive and each segment must start with a letter or underscore.
            var stem = CompactName(string.Concat(id.Select(c => char.IsAsciiLetterLower(c) || char.IsAsciiDigit(c) ? c.ToString() : $"_{(int)c:X2}")));
            if (char.IsAsciiDigit(stem[0])) stem = "_" + stem;
            var name = "JS." + stem + ".DOUBLE";
            sections.Add($$"""
                // Apply scripts, then try ={{name}}(21) in a cell.
                // UDFs use only their arguments; workbook access belongs in Actions.
                // Register events in the main entry or a Common module imported by it.
                {{(typed ? "" : "/** @param {unknown} value @returns {number} */\n")}}function double(value{{(typed ? ": unknown" : "")}}){{(typed ? ": number" : "")}} {
                    if (typeof value !== 'number' || !Number.isFinite(value)) {
                        throw new TypeError('Expected a finite number.');
                    }
                    return value * 2;
                }

                host.functions.register({
                    name: '{{name}}',
                    apiVersion: 2,
                    parameters: [{ name: 'value', type: 'any' }],
                    returns: 'number'
                }, double);
                """);
        }
        if (actions)
            sections.Add($$"""
                // Apply scripts, select this module in the action picker, then run it.
                // One file defines one action through main; use a separate module for another action.
                // Register events in the main entry or a Common module imported by it.
                // Writes commit together and can be undone as one operation.
                export function main(workbook{{(typed ? ": QuickExcel.Workbook" : "")}}) {
                    const cell = workbook.getActiveCell();
                    cell.setValue(42);
                    cell.getFormat().getFill().setColor('#FFF2CC');
                    console.log('Updated', cell.getAddress());
                }
                """);
        if (sections.Count == 0)
            sections.Add($$"""
                // Import this Common module by its logical module ID: '{{id}}'.
                // To register events here, import this module from the main entry.
                export function example(value{{(typed ? ": number" : "")}}) {
                    return value;
                }
                """);
        return string.Join("\n\n", sections) + "\n";
    }

    private static string CompactName(string value) => value.Length <= 80 ? value :
        value[..60] + "_" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)))[..16];
}
