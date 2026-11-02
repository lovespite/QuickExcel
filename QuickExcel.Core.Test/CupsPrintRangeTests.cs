using QuickExcel.Printing;

namespace QuickExcel.Core.Test;
public sealed class CupsPrintRangeTests
{
    [Theory]
    [InlineData("",7,"1-7")]
    [InlineData("3-5,1,2,4,7",7,"1-5,7")]
    [InlineData("2",7,"2")]
    public void PreservesSelectedPagesWithoutDuplicates(string input,int pages,string expected)
        =>Assert.Equal(expected,CupsPrintService.NormalizePageRanges(input,pages));
    [Theory]
    [InlineData("0")][InlineData("8")][InlineData("5-2")][InlineData("1-3-5")][InlineData("1,,3")][InlineData("-o raw")]
    public void RejectsInvalidAndOutsidePages(string input)
        =>Assert.Throws<ArgumentException>(()=>CupsPrintService.NormalizePageRanges(input,7));
    [Fact]public void NeverEnumeratesUnboundedPageDomain()=>Assert.Throws<ArgumentException>(()=>CupsPrintService.NormalizePageRanges("1-2147483647",int.MaxValue));
}
