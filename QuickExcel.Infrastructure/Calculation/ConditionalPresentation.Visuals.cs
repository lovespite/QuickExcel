using System.Globalization;
using QuickExcel.Domain;

namespace QuickExcel.Calculation;

public static partial class ConditionalPresentation
{
    private sealed partial class Cache
    {
        private readonly Dictionary<(SheetId, string), double[]> _numbers = [];
        private long _numericCells;
        private double[] Numbers(WorksheetSnapshot sheet, ConditionalFormatRule rule)
        {
            var key = (sheet.Id, string.Join(' ', rule.Ranges));
            if (_numbers.TryGetValue(key, out var cached)) return cached;
            var values = new List<double>();
            // Sparse enumeration also permits whole-column quick formatting without allocating blank cells.
            foreach (var position in sheet.Cells.Keys.Where(p => rule.Contains(p.Row, p.Column)))
            {
                if (++_numericCells > 2_000_000) throw new NotSupportedException("视觉条件格式累计数据超过 2,000,000 格，请缩小规则范围。");
                var value = _evaluation.Cell(new(sheet.Id, position.Row, position.Column));
                if (value.Kind == CellValueKind.Number) values.Add(value.NumberValue);
            }
            values.Sort(); return _numbers[key] = values.ToArray();
        }
        private double Threshold(ConditionalThreshold threshold, double[] numbers, WorksheetSnapshot sheet, ConditionalFormatRule rule)
        {
            var min = numbers[0]; var max = numbers[^1];
            return threshold.Kind switch
            {
                ConditionalThresholdKind.Minimum => min, ConditionalThresholdKind.Maximum => max,
                ConditionalThresholdKind.AutoMinimum => Math.Min(0, min), ConditionalThresholdKind.AutoMaximum => Math.Max(0, max),
                ConditionalThresholdKind.Number => threshold.Value,
                ConditionalThresholdKind.Percent => Interpolate(min, max, threshold.Value / 100),
                ConditionalThresholdKind.Percentile => Percentile(numbers, threshold.Value / 100),
                ConditionalThresholdKind.Formula => _evaluation.Scalar(threshold.Formula!, sheet.Id, rule.Anchor, rule.Anchor) is { Kind: CellValueKind.Number } value ? value.NumberValue : double.NaN,
                _ => double.NaN
            };
        }
        private static double Percentile(double[] numbers, double p)
        {
            var rank = (numbers.Length - 1) * p; var low = (int)Math.Floor(rank); var high = (int)Math.Ceiling(rank);
            return Interpolate(numbers[low], numbers[high], rank - low);
        }
        private static double Interpolate(double low, double high, double fraction) => low * (1 - fraction) + high * fraction;
        private static double Fraction(double value, double low, double high)
        {
            if (low == high) return 1;
            // Scaling avoids overflow for finite values near Double.MaxValue.
            var scale = Math.Max(Math.Abs(low), Math.Abs(high));
            return Math.Clamp((value / scale - low / scale) / (high / scale - low / scale), 0, 1);
        }
        private static string Blend(string low, string high, double fraction)
        {
            var a = uint.Parse(low, NumberStyles.HexNumber); var b = uint.Parse(high, NumberStyles.HexNumber); uint result = 0;
            for (var shift = 0; shift <= 24; shift += 8)
                result |= (uint)Math.Round(Interpolate((a >> shift) & 255, (b >> shift) & 255, fraction)) << shift;
            return result.ToString("X8", CultureInfo.InvariantCulture);
        }
        private bool VisualMatch(WorksheetSnapshot sheet, CellPosition position, ConditionalFormatRule rule, out DifferentialStyle style, out ConditionalCellVisual? visual)
        {
            style = rule.Style; visual = null;
            var value = _evaluation.Cell(new(sheet.Id, position.Row, position.Column));
            if (value.Kind != CellValueKind.Number) return false;
            var numbers = Numbers(sheet, rule); if (numbers.Length == 0) return false;
            var thresholds = rule.Thresholds.Select(t => Threshold(t, numbers, sheet, rule)).ToArray();
            if (thresholds.Any(t => !double.IsFinite(t))) return false;
            if (thresholds.Zip(thresholds.Skip(1)).Any(p => p.First > p.Second))
                throw new NotSupportedException("条件格式阈值须从小到大排列。");
            var n = value.NumberValue;
            if (rule.ColorScale is { } scale)
            {
                var segment = thresholds.Length == 3 && n > thresholds[1] ? 1 : 0;
                style = style with { FillColor = Blend(scale.Colors[segment], scale.Colors[segment + 1], Fraction(n, thresholds[segment], thresholds[segment + 1])) };
            }
            else if (rule.IconSet is { } icons)
            {
                var index = 0;
                for (var i = 1; i < thresholds.Length; i++) if (n > thresholds[i] || icons.Thresholds[i].Inclusive && n == thresholds[i]) index = i;
                if (icons.Reverse) index = thresholds.Length - 1 - index;
                visual = new(Icon: new(icons.Name, index), HideValue: !icons.ShowValue);
            }
            else if (rule.DataBar is { } bar)
            {
                var low = thresholds[0]; var high = thresholds[1]; double start, end; double? axis = null;
                var minimum = bar.MinLength / 100d; var maximum = bar.MaxLength / 100d;
                if (bar.Axis == DataBarAxis.None)
                { start = 0; end = Interpolate(minimum, maximum, Fraction(n, low, high)); }
                else
                {
                    axis = bar.Axis == DataBarAxis.Middle ? .5 : low >= 0 ? 0 : high <= 0 ? 1 : Fraction(0, low, high);
                    start = axis.Value;
                    var ratio = low >= 0 ? Fraction(n, low, high) : high <= 0 ? low == high ? 1 : 1 - Fraction(n, low, high) :
                        n < 0 ? Math.Clamp(n / low, 0, 1) : Math.Clamp(n / high, 0, 1);
                    var length = n == 0 ? 0 : Interpolate(minimum, maximum, ratio);
                    end = start + (n < 0 ? -start : 1 - start) * length;
                }
                if (bar.Direction == DataBarDirection.RightToLeft)
                { start = 1 - start; end = 1 - end; if (axis is { } at) axis = 1 - at; }
                visual = new(new(start, end, axis, n < 0 ? bar.NegativeColor : bar.Color, bar.Gradient,
                    n < 0 ? bar.NegativeBorderColor ?? bar.BorderColor : bar.BorderColor, bar.AxisColor), HideValue: !bar.ShowValue);
            }
            return true;
        }
    }
}
