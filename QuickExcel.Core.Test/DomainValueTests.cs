using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class DomainValueTests
{
    [Fact]
    public void BlankEmptyTextAndLookalikeValuesRemainDistinct()
    {
        Assert.Equal(CellValue.Blank, default(CellValue));
        Assert.Equal(CellValueKind.Blank, default(CellValue).Kind);

        CellValue[] values =
        [
            CellValue.Blank,
            CellValue.FromText(""),
            CellValue.FromText("00123"),
            CellValue.FromNumber(123),
            CellValue.FromText("TRUE"),
            CellValue.FromBoolean(true),
            CellValue.FromNumber(1),
            CellValue.FromText("#VALUE!"),
            CellValue.FromError(CellErrorCode.Value)
        ];

        Assert.Equal(values.Length, values.Distinct().Count());
    }

    [Theory]
    [InlineData("00123")]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("#VALUE!")]
    [InlineData("TRUE")]
    [InlineData("=1+1")]
    public void TextRetainsItsExactInputWithoutParsing(string text)
    {
        var value = CellValue.FromText(text);

        Assert.Equal(CellValueKind.Text, value.Kind);
        Assert.Equal(text, value.TextValue);
    }

    [Fact]
    public void NullTextIsRejectedInsteadOfBecomingBlank() =>
        Assert.Throws<ArgumentNullException>(() => CellValue.FromText(null!));

    [Theory]
    [InlineData(0d)]
    [InlineData(-1.25d)]
    [InlineData(double.Epsilon)]
    [InlineData(double.MaxValue)]
    [InlineData(double.MinValue)]
    public void FiniteNumbersRetainTheirNumericValue(double number)
    {
        var value = CellValue.FromNumber(number);

        Assert.Equal(CellValueKind.Number, value.Kind);
        Assert.Equal(number, value.NumberValue);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteNumbersBecomeNumberErrors(double number)
    {
        var value = CellValue.FromNumber(number);

        Assert.Equal(CellValueKind.Error, value.Kind);
        Assert.Equal(CellErrorCode.Number, value.ErrorCode);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BooleansRemainTypedBooleans(bool boolean)
    {
        var value = CellValue.FromBoolean(boolean);

        Assert.Equal(CellValueKind.Boolean, value.Kind);
        Assert.Equal(boolean, value.BooleanValue);
        Assert.NotEqual(CellValue.FromNumber(boolean ? 1 : 0), value);
    }

    [Fact]
    public void ErrorCodesRemainDistinctTypedValues()
    {
        var errors = Enum.GetValues<CellErrorCode>();
        var values = errors.Select(CellValue.FromError).ToArray();

        Assert.Equal(errors.Length, values.Distinct().Count());
        for (var index = 0; index < errors.Length; index++)
        {
            Assert.Equal(CellValueKind.Error, values[index].Kind);
            Assert.Equal(errors[index], values[index].ErrorCode);
        }
    }

    [Fact]
    public void TypedAccessorsRejectImplicitConversions()
    {
        CellValue[] values =
        [
            CellValue.Blank,
            CellValue.FromText("123"),
            CellValue.FromNumber(123),
            CellValue.FromBoolean(true),
            CellValue.FromError(CellErrorCode.Value)
        ];

        foreach (var value in values)
        {
            if (value.Kind != CellValueKind.Text)
                Assert.Throws<InvalidOperationException>(() => value.TextValue);
            if (value.Kind != CellValueKind.Number)
                Assert.Throws<InvalidOperationException>(() => value.NumberValue);
            if (value.Kind != CellValueKind.Boolean)
                Assert.Throws<InvalidOperationException>(() => value.BooleanValue);
            if (value.Kind != CellValueKind.Error)
                Assert.Throws<InvalidOperationException>(() => value.ErrorCode);
        }
    }

    [Fact]
    public void DiagnosticFormattingDoesNotReadInactiveTypedPayloads()
    {
        CellValue[] values =
        [
            CellValue.Blank,
            CellValue.FromText("sample"),
            CellValue.FromNumber(42),
            CellValue.FromBoolean(true),
            CellValue.FromError(CellErrorCode.Value)
        ];

        foreach (var value in values)
        {
            Assert.False(string.IsNullOrWhiteSpace(value.ToString()));
            Assert.False(string.IsNullOrWhiteSpace(CellContent.FromLiteral(value).ToString()));
        }
        Assert.False(string.IsNullOrWhiteSpace(CellContent.FromFormula("=1+2").ToString()));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(int.MaxValue)]
    public void UndefinedErrorCodesAreRejected(int error) =>
        Assert.ThrowsAny<ArgumentException>(() => CellValue.FromError((CellErrorCode)error));

    [Fact]
    public void LiteralFormulaTextIsNotReinterpretedAsAFormula()
    {
        var literal = CellContent.FromLiteral(CellValue.FromText("=SUM("), 17);

        Assert.Equal(CellContentKind.Literal, literal.Kind);
        Assert.Equal(CellValue.FromText("=SUM("), literal.LiteralValue);
        Assert.Equal(17, literal.StyleId);
        Assert.Throws<InvalidOperationException>(() => literal.FormulaText);
    }

    [Theory]
    [InlineData("=1+2")]
    [InlineData("= SUM( A1 , B1 ) \r\n")]
    [InlineData("=SUM(")]
    [InlineData("=)")]
    public void FormulaPreservesInputIncludingUnparsedInvalidSyntax(string formula)
    {
        var content = CellContent.FromFormula(formula, 8);

        Assert.Equal(CellContentKind.Formula, content.Kind);
        Assert.Equal(formula, content.FormulaText);
        Assert.Equal(8, content.StyleId);
        Assert.Throws<InvalidOperationException>(() => content.LiteralValue);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("=")]
    [InlineData("= \t\r\n")]
    [InlineData("SUM(A1)")]
    [InlineData(" =1")]
    public void FormulaRequiresAnEqualsSignAndNonBlankBody(string? formula) =>
        Assert.ThrowsAny<ArgumentException>(() => CellContent.FromFormula(formula!));

    [Fact]
    public void ContentDefaultsToStyleZeroAndRejectsNegativeStyleIds()
    {
        Assert.Equal(0, CellContent.FromLiteral(CellValue.Blank).StyleId);
        Assert.Equal(0, CellContent.FromFormula("=1").StyleId);
        Assert.ThrowsAny<ArgumentException>(() => CellContent.FromLiteral(CellValue.Blank, -1));
        Assert.ThrowsAny<ArgumentException>(() => CellContent.FromFormula("=1", -1));
    }

    [Fact]
    public void CalculationResultsDoNotReplaceFormulaInput()
    {
        var content = CellContent.FromFormula("=1+2", 5);
        var first = new CellCalculation(CellValue.FromNumber(3), CalculationState.Fresh, 12, 3, 7);
        var pending = new CellCalculation(first.Value, CalculationState.Stale, 13, 4, 8);

        Assert.Equal("=1+2", content.FormulaText);
        Assert.Equal(5, content.StyleId);
        Assert.Equal(CellValue.FromNumber(3), first.Value);
        Assert.Equal(CalculationState.Fresh, first.State);
        Assert.Equal(12, first.InputRevision);
        Assert.Equal(3, first.CalculationGeneration);
        Assert.Equal(7, first.ScriptGeneration);
        Assert.Equal(first.Value, pending.Value);
        Assert.Equal(CalculationState.Stale, pending.State);
    }

    [Fact]
    public void MissingCacheAndCachedBlankAreDifferent()
    {
        var missing = new CellCalculation(null, CalculationState.Dirty, 0, 0, 0);
        var blank = new CellCalculation(CellValue.Blank, CalculationState.Dirty, 0, 0, 0);

        Assert.Null(missing.Value);
        Assert.True(blank.Value.HasValue);
        Assert.Equal(CellValue.Blank, blank.Value.Value);
    }

    [Theory]
    [InlineData(CalculationState.Dirty)]
    [InlineData(CalculationState.Running)]
    [InlineData(CalculationState.Stale)]
    [InlineData(CalculationState.Blocked)]
    public void PendingStatesCanRetainAnOldErrorOrHaveNoCache(CalculationState state)
    {
        CellValue?[] cachedValues = [null, CellValue.Blank, CellValue.FromNumber(42), CellValue.FromError(CellErrorCode.Script)];

        foreach (var cached in cachedValues)
        {
            var calculation = new CellCalculation(cached, state, 4, 2, 1);
            Assert.Equal(state, calculation.State);
            Assert.Equal(cached, calculation.Value);
        }
    }

    [Fact]
    public void FreshAndErrorStatesRequireMatchingCachedValues()
    {
        CellValue[] successfulValues = [CellValue.Blank, CellValue.FromText(""), CellValue.FromNumber(0), CellValue.FromBoolean(false)];
        foreach (var value in successfulValues)
        {
            Assert.Equal(value, new CellCalculation(value, CalculationState.Fresh, 0, 0, 0).Value);
            Assert.ThrowsAny<ArgumentException>(() => new CellCalculation(value, CalculationState.Error, 0, 0, 0));
        }

        var error = CellValue.FromError(CellErrorCode.Cycle);
        Assert.Equal(error, new CellCalculation(error, CalculationState.Error, 0, 0, 0).Value);
        Assert.ThrowsAny<ArgumentException>(() => new CellCalculation(error, CalculationState.Fresh, 0, 0, 0));
        Assert.ThrowsAny<ArgumentException>(() => new CellCalculation(null, CalculationState.Fresh, 0, 0, 0));
        Assert.ThrowsAny<ArgumentException>(() => new CellCalculation(null, CalculationState.Error, 0, 0, 0));
    }

    [Theory]
    [InlineData(-1, 0, 0)]
    [InlineData(0, -1, 0)]
    [InlineData(0, 0, -1)]
    public void NegativeCalculationVersionsAreRejected(long inputRevision, long calculationGeneration, long scriptGeneration) =>
        Assert.ThrowsAny<ArgumentException>(() => new CellCalculation(null, CalculationState.Dirty, inputRevision, calculationGeneration, scriptGeneration));

    [Theory]
    [InlineData(-1)]
    [InlineData(int.MaxValue)]
    public void UndefinedCalculationStatesAreRejected(int state) =>
        Assert.ThrowsAny<ArgumentException>(() => new CellCalculation(null, (CalculationState)state, 0, 0, 0));

    [Fact]
    public void DiagnosticsAreCopiedAndCannotBeChangedThroughTheExposedCollection()
    {
        var diagnostics = new List<string> { "Unknown function", "A1 depends on B1" };
        var calculation = new CellCalculation(null, CalculationState.Blocked, 1, 2, 3, diagnostics);
        diagnostics[0] = "Changed after calculation";
        diagnostics.Clear();

        Assert.Equal(new[] { "Unknown function", "A1 depends on B1" }, calculation.Diagnostics);
        if (calculation.Diagnostics is IList<string> mutableView)
        {
            Assert.True(mutableView.IsReadOnly);
            Assert.Throws<NotSupportedException>(() => mutableView[0] = "Changed through the result");
        }
        Assert.Empty(new CellCalculation(null, CalculationState.Dirty, 0, 0, 0).Diagnostics);
        Assert.ThrowsAny<ArgumentException>(() =>
            new CellCalculation(null, CalculationState.Dirty, 0, 0, 0, new[] { "Valid diagnostic", null! }));
    }

    [Theory]
    [InlineData(0, 1899, 12, 31)]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(61, 1900, 3, 1)]
    public void WindowsDatesRespectTheHistoricalLeapDayBoundary(double serial, int year, int month, int day)
    {
        var date = new ExcelSerialDate(serial, ExcelDateSystem.Windows1900);

        Assert.False(date.IsFictitiousLeapDay);
        Assert.True(date.TryToDateTime(out var actual));
        Assert.Equal(new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Unspecified), actual);
        Assert.Equal(DateTimeKind.Unspecified, actual.Kind);
        Assert.Equal(date, ExcelSerialDate.FromDateTime(actual, ExcelDateSystem.Windows1900));
    }

    [Theory]
    [InlineData(60d)]
    [InlineData(60.5d)]
    [InlineData(60.999d)]
    public void FictitiousWindowsLeapDayIsPreservedButCannotBecomeARealDate(double serial)
    {
        var date = new ExcelSerialDate(serial, ExcelDateSystem.Windows1900);

        Assert.Equal(serial, date.Serial);
        Assert.Equal(ExcelDateSystem.Windows1900, date.DateSystem);
        Assert.True(date.IsFictitiousLeapDay);
        Assert.False(date.TryToDateTime(out _));
        Assert.False(date.TryConvertTo(ExcelDateSystem.Mac1904, out _));
        Assert.True(date.TryConvertTo(ExcelDateSystem.Windows1900, out var unchanged));
        Assert.Equal(date, unchanged);
    }

    [Theory]
    [InlineData(0, 1904, 1, 1)]
    [InlineData(1, 1904, 1, 2)]
    [InlineData(59, 1904, 2, 29)]
    [InlineData(60, 1904, 3, 1)]
    public void MacDatesUseTheirOwnEpochAndRealLeapYear(double serial, int year, int month, int day)
    {
        var date = new ExcelSerialDate(serial, ExcelDateSystem.Mac1904);

        Assert.False(date.IsFictitiousLeapDay);
        Assert.True(date.TryToDateTime(out var actual));
        Assert.Equal(new DateTime(year, month, day), actual);
        Assert.Equal(DateTimeKind.Unspecified, actual.Kind);
        Assert.Equal(date, ExcelSerialDate.FromDateTime(actual, ExcelDateSystem.Mac1904));
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900, -0.5, 1899, 12, 30, 12)]
    [InlineData(ExcelDateSystem.Windows1900, 0.25, 1899, 12, 31, 6)]
    [InlineData(ExcelDateSystem.Windows1900, 59.5, 1900, 2, 28, 12)]
    [InlineData(ExcelDateSystem.Windows1900, 61.75, 1900, 3, 1, 18)]
    [InlineData(ExcelDateSystem.Mac1904, -0.5, 1903, 12, 31, 12)]
    [InlineData(ExcelDateSystem.Mac1904, 0.25, 1904, 1, 1, 6)]
    public void NegativeAndFractionalSerialsRetainTheirTimeOfDay(ExcelDateSystem system, double serial, int year, int month, int day, int hour)
    {
        var date = new ExcelSerialDate(serial, system);
        var expected = new DateTime(year, month, day, hour, 0, 0, DateTimeKind.Unspecified);

        Assert.True(date.TryToDateTime(out var actual));
        Assert.Equal(expected, actual);
        Assert.Equal(date, ExcelSerialDate.FromDateTime(expected, system));
    }

    [Theory]
    [InlineData(1462d, 0d)]
    [InlineData(1462.5d, 0.5d)]
    [InlineData(45000.25d, 43538.25d)]
    [InlineData(61d, -1401d)]
    [InlineData(59d, -1402d)]
    [InlineData(0d, -1461d)]
    public void DateSystemConversionPreservesRealDatesOnBothSidesOfTheLeapDay(double windowsSerial, double macSerial)
    {
        var windows = new ExcelSerialDate(windowsSerial, ExcelDateSystem.Windows1900);
        var mac = new ExcelSerialDate(macSerial, ExcelDateSystem.Mac1904);

        Assert.True(windows.TryConvertTo(ExcelDateSystem.Mac1904, out var actualMac));
        Assert.Equal(mac, actualMac);
        Assert.True(mac.TryConvertTo(ExcelDateSystem.Windows1900, out var actualWindows));
        Assert.Equal(windows, actualWindows);
    }

    [Theory]
    [InlineData(DateTimeKind.Local)]
    [InlineData(DateTimeKind.Utc)]
    public void DateConversionRejectsTimezoneBearingValuesWithoutImplicitConversion(DateTimeKind kind)
    {
        var date = new DateTime(2026, 10, 22, 8, 30, 0, kind);

        Assert.ThrowsAny<ArgumentException>(() => ExcelSerialDate.FromDateTime(date, ExcelDateSystem.Windows1900));
        Assert.ThrowsAny<ArgumentException>(() => ExcelSerialDate.FromDateTime(date, ExcelDateSystem.Mac1904));
    }

    [Theory]
    [InlineData(double.MaxValue)]
    [InlineData(double.MinValue)]
    public void HugeFiniteSerialsRemainRepresentableButCannotBecomeDateTime(double serial)
    {
        foreach (var system in Enum.GetValues<ExcelDateSystem>())
        {
            var date = new ExcelSerialDate(serial, system);
            Assert.Equal(serial, date.Serial);
            Assert.False(date.TryToDateTime(out _));
            Assert.True(date.TryConvertTo(system, out var unchanged));
            Assert.Equal(date, unchanged);
            var other = system == ExcelDateSystem.Windows1900 ? ExcelDateSystem.Mac1904 : ExcelDateSystem.Windows1900;
            Assert.False(date.TryConvertTo(other, out _));
        }
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteDateSerialsAreRejected(double serial) =>
        Assert.ThrowsAny<ArgumentException>(() => new ExcelSerialDate(serial, ExcelDateSystem.Windows1900));

    [Fact]
    public void DateTimeRangeEdgesDoNotSilentlyProduceUnconvertibleDates()
    {
        foreach (var system in Enum.GetValues<ExcelDateSystem>())
        {
            var earliest = ExcelSerialDate.FromDateTime(DateTime.MinValue, system);
            Assert.True(earliest.TryToDateTime(out var actual));
            Assert.Equal(DateTime.MinValue, actual);
            // At this magnitude double rounds the final tick up to year 10000.
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelSerialDate.FromDateTime(DateTime.MaxValue, system));
        }
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(int.MaxValue)]
    public void UndefinedDateSystemsAreRejectedAtEveryEntryPoint(int system)
    {
        var invalid = (ExcelDateSystem)system;
        var date = new ExcelSerialDate(1, ExcelDateSystem.Windows1900);

        Assert.ThrowsAny<ArgumentException>(() => new ExcelSerialDate(1, invalid));
        Assert.ThrowsAny<ArgumentException>(() => ExcelSerialDate.FromDateTime(new DateTime(2026, 10, 22), invalid));
        Assert.ThrowsAny<ArgumentException>(() => date.TryConvertTo(invalid, out _));
    }
}
