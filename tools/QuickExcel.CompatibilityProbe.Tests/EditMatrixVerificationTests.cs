using System.Text.Json;
using QuickExcel.CompatibilityProbe;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed class EditMatrixVerificationTests
{
    [Theory]
    [InlineData("=SUM('数据'!A2:B2)", "=SUM(数据!A2:B2)", true)]
    [InlineData("=A1*(B1+2)", "=(A1*(B1+2))", true)]
    [InlineData("='数据'!#REF!*2", "=(#REF!*2)", true)]
    [InlineData("=A1*(B1+2)", "=A1*B1+2", false)]
    [InlineData("=\"'数据'!A2\"", "=\"数据!A2\"", false)]
    [InlineData("=A1", "=$A$1", false)]
    [InlineData("=Sheet1!A1", "=Sheet2!A1", false)]
    [InlineData("=Table1[列]", "=Table2[列]", false)]
    public void FormulaEquivalenceIsSyntacticAndDoesNotMaskChangedOperandsStringsOrReferences(string expected, string actual, bool equivalent)
    {
        using var a = JsonDocument.Parse(JsonSerializer.Serialize(new { Formula = expected }));
        using var b = JsonDocument.Parse(JsonSerializer.Serialize(new { Formula = actual }));
        var diff = Assert.Single(EditMatrixVerification.Compare(a.RootElement, b.RootElement));
        Assert.Equal(equivalent ? "equivalent-formula-serialization" : "unexpected", diff.Classification);
    }
    [Theory]
    [InlineData("{\"NumberFormat\":\"0.00\"}", "{\"NumberFormat\":\"General\"}")]
    [InlineData("{\"PrintArea\":\"$A$1:$D$6\"}", "{\"PrintArea\":\"\"}")]
    [InlineData("{\"Cells\":[{\"Value\":\"00123\"}]}", "{\"Cells\":[{\"Value\":123}]}")]
    [InlineData("{\"Cells\":[1,2]}", "{\"Cells\":[1]}")]
    [InlineData("{\"Font\":\"Arial\"}", "{}")]
    public void LostStyleMetadataTypeCellsOrFieldsFail(string expected, string actual)
    {
        using var a = JsonDocument.Parse(expected); using var b = JsonDocument.Parse(actual);
        Assert.Contains(EditMatrixVerification.Compare(a.RootElement, b.RootElement), d => d.Classification == "unexpected");
    }
}
