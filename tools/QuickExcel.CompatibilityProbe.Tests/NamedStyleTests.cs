using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed partial class CompatibilityProbeTests
{
    [Fact]
    public void QuotationApplyFlagsAndFontSchemeNowSurviveAllEditPathsWithoutLoss()
    {
        var report = ProbeRunner.Run(Fixture("quotation"), NewEvidenceDirectory());
        Assert.True(report.SourceUnchanged);
        Assert.All(report.Paths, path =>
        {
            Assert.Empty(path.CompatibilityIssues);
            Assert.NotNull(path.Output); Assert.Empty(path.Output.Issues);
            Assert.DoesNotContain(path.Diagnostics, d => d.Feature == "styles/inheritance");
        });
        Assert.Equal("saved-modeled-subset", report.Paths.Single(p => p.Mode == "style").Status);
    }
}
