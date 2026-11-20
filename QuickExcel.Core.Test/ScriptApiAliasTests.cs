using QuickExcel.Scripting;

namespace QuickExcel.Core.Test;

/// <summary>The legacy $ / quickExcel globals and snapshot/patch context are gone.</summary>
public sealed class ScriptApiAliasTests
{
    [Fact]
    public void LegacyGlobalsDoNotExistAndHostRegistersFunctions()
    {
        using var session = new JscSession("""
            if (typeof $ !== 'undefined' || typeof quickExcel !== 'undefined') throw Error('legacy global');
            if (typeof host !== 'object' || !host.functions || !host.events) throw Error('host');
            host.functions.register('OK', () => 1);
            """);
        Assert.Equal(1d, session.Invoke("OK", []));
    }

    [Fact]
    public void MainReceivesOnlyWorkbookAndHostWithoutSnapshotOrPatch()
    {
        using var session = new JscSession(new Dictionary<string, string>
        {
            ["main"] = """
                export function main(workbook, host) {
                  if (typeof $ !== 'undefined' || typeof quickExcel !== 'undefined') throw Error('legacy global');
                  if (typeof workbook !== 'object' || typeof host !== 'object') throw Error('arguments');
                  if (typeof snapshot !== 'undefined' || typeof patch !== 'undefined') throw Error('legacy context');
                  console.log('ran');
                }
                """
        }, ["main"], actions: true);
        Assert.Equal("main", Assert.Single(session.ActionDescriptors).Name);
        session.HostRequest = json =>
            json.Contains("\"files\"", StringComparison.Ordinal) ? "{\"ok\":true,\"value\":[]}" : "{\"ok\":true,\"value\":null}";
        session.RunAction("main", System.Text.Json.JsonSerializer.SerializeToElement(new { operationId = Guid.NewGuid(), name = "main", snapshot = new { baseRevision = 0 }, liveModel = true }));
        Assert.Contains("ran", session.GetLogs());
    }

    [Fact]
    public void ModuleDescriptionSurfacesOnActionDescriptorWithoutAffectingExecution()
    {
        using var session = new JscSession(new Dictionary<string, string>
        {
            ["main"] = """
                export function main(workbook, host) { console.log('ran'); }
                """,
            ["helper"] = "export const value = 1;"
        }, ["main"], actions: true, descriptions: new Dictionary<string, string> { ["main"] = "填充示例单元格", ["helper"] = "不应显示在操作列表" });
        var action = Assert.Single(session.ActionDescriptors);
        Assert.Equal("main", action.Name);
        Assert.Equal("填充示例单元格", action.Description);
        session.HostRequest = json => "{\"ok\":true,\"value\":null}";
        session.RunAction("main", System.Text.Json.JsonSerializer.SerializeToElement(new { operationId = Guid.NewGuid(), name = "main", snapshot = new { baseRevision = 0 }, liveModel = true }));
        Assert.Contains("ran", session.GetLogs());
    }
}
