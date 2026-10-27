using QuickExcel.Calculation;
using QuickExcel.Runtime;
using QuickExcel.Core;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

/// <summary>The only desktop location that chooses concrete IO, calculation and native runtime adapters.</summary>
internal static class DesktopComposition
{
    public static WorkbookServices CreateServices() =>
        new(new LegacyWorkbookStorage(), new LegacyWorkbookCalculator(new QuickJsRuntimeFactory()));
}

