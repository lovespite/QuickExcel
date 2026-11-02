using System.ComponentModel;
using System.Runtime.InteropServices;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Windows common print dialog and bounded GDI output. No WPF/System.Drawing or browser printing.</summary>
internal sealed class WindowsPrintJob : IDisposable
{
    private nint _dc;
    private readonly int[] _pages;
    private readonly PrinterRasterLayout _layout;
    private WindowsPrintJob(nint dc, int[] pages, PrinterRasterLayout layout) { _dc = dc; _pages = pages; _layout = layout; }
    internal static WindowsPrintJob? Choose(nint owner, PrintSession session)
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
        if (owner == 0) throw new ArgumentException("系统打印对话框需要所属窗口。");
        var ranges = Marshal.AllocHGlobal(16 * Marshal.SizeOf<PageRange>());
        var dialog = new PrintDialog { Size = (uint)Marshal.SizeOf<PrintDialog>(), Owner = owner, Flags = ReturnDc | UseDriverCopies | NoSelection | NoCurrentPage | HidePrintToFile | DisablePrintToFile,
            MaxRanges = 16, Ranges = ranges, MinPage = 1, MaxPage = (uint)session.Job.Layout.Pages.Length, Copies = 1, StartPage = uint.MaxValue };
        try
        {
            // Retrieve default device settings so the native dialog initially reflects this immutable job.
            dialog.Flags |= ReturnDefault;
            var initial = PrintDlgExW(ref dialog);
            if (dialog.Dc != 0) { DeleteDC(dialog.Dc); dialog.Dc = 0; }
            if (initial >= 0 && dialog.DevMode != 0) ConfigureDefault(dialog.DevMode, session);
            dialog.Flags &= ~ReturnDefault; dialog.ResultAction = 0;
            var result = PrintDlgExW(ref dialog); if (result < 0) Marshal.ThrowExceptionForHR(result);
            if (dialog.ResultAction != 1) return null;
            if (dialog.Dc == 0) throw new IOException("打印机没有提供绘图上下文。");
            var selected = new SortedSet<int>();
            if ((dialog.Flags & PageNumbers) != 0)
            {
                if (dialog.RangeCount is < 1 or > 16) throw new InvalidDataException("打印页范围无效。");
                for (var i = 0; i < dialog.RangeCount; i++)
                {
                    var range = Marshal.PtrToStructure<PageRange>(ranges + i * Marshal.SizeOf<PageRange>());
                    if (range.From < 1 || range.To < range.From || range.To > session.Job.Layout.Pages.Length) throw new InvalidDataException("打印页范围超出当前预览。");
                    for (var page = range.From; page <= range.To; page++) selected.Add((int)page - 1);
                }
            }
            else for (var page = 0; page < session.Job.Layout.Pages.Length; page++) selected.Add(page);
            var dpiX = GetDeviceCaps(dialog.Dc, 88); var dpiY = GetDeviceCaps(dialog.Dc, 90);
            var width = GetDeviceCaps(dialog.Dc, 110); var height = GetDeviceCaps(dialog.Dc, 111);
            var offsetX = GetDeviceCaps(dialog.Dc, 112); var offsetY = GetDeviceCaps(dialog.Dc, 113);
            var layout = PrinterRasterLayout.Create(session.Job.Layout.PaperWidth, session.Job.Layout.PaperHeight, dpiX, dpiY, width, height, offsetX, offsetY);
            var margins = session.Job.Layout.Margins;
            var right = width - offsetX - GetDeviceCaps(dialog.Dc, 8); var bottom = height - offsetY - GetDeviceCaps(dialog.Dc, 10);
            if (margins.Left * dpiX + 1 < offsetX || margins.Right * dpiX + 1 < right || margins.Top * dpiY + 1 < offsetY || margins.Bottom * dpiY + 1 < bottom)
                throw new InvalidOperationException("页面边距小于此打印机的不可打印边缘。请增加页面边距并重新生成预览。");
            var dc = dialog.Dc; dialog.Dc = 0; return new(dc, selected.ToArray(), layout);
        }
        finally
        {
            if (dialog.Dc != 0) DeleteDC(dialog.Dc);
            if (dialog.DevMode != 0) GlobalFree(dialog.DevMode); if (dialog.DevNames != 0) GlobalFree(dialog.DevNames);
            Marshal.FreeHGlobal(ranges);
        }
    }
    private static void ConfigureDefault(nint handle, PrintSession session)
    {
        var pointer = GlobalLock(handle); if (pointer == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
        try
        {
            // Public DEVMODEW printer header offsets, preserving the driver's variable-sized private tail.
            if (GlobalSize(handle) < 94 || (ushort)Marshal.ReadInt16(pointer, 68) < 94) throw new InvalidDataException("打印机 DEVMODEW 过短。");
            const int orientation = 1, paperSize = 2, paperLength = 4, paperWidth = 8, scale = 16, copies = 256, color = 2048;
            var setup = session.Job.Sheet.Printing.Setup; var landscape = session.Job.Layout.PaperWidth > session.Job.Layout.PaperHeight;
            var fields = Marshal.ReadInt32(pointer, 72) | orientation | scale | copies | color;
            Marshal.WriteInt16(pointer, 76, (short)(landscape ? 2 : 1)); Marshal.WriteInt16(pointer, 84, 100);
            Marshal.WriteInt16(pointer, 86, checked((short)Math.Clamp(setup?.Copies ?? 1, 1, 999))); Marshal.WriteInt16(pointer, 92, (short)(setup?.BlackAndWhite == true ? 1 : 2));
            if (setup?.PaperWidth is not null || setup?.PaperHeight is not null)
            {
                fields = (fields & ~paperSize) | paperLength | paperWidth;
                var width = landscape ? session.Job.Layout.PaperHeight : session.Job.Layout.PaperWidth;
                var height = landscape ? session.Job.Layout.PaperWidth : session.Job.Layout.PaperHeight;
                Marshal.WriteInt16(pointer, 80, checked((short)Math.Round(height * 254 / 96))); Marshal.WriteInt16(pointer, 82, checked((short)Math.Round(width * 254 / 96)));
            }
            else { fields = (fields & ~(paperLength | paperWidth)) | paperSize; Marshal.WriteInt16(pointer, 78, checked((short)(setup?.PaperSize ?? 9))); }
            Marshal.WriteInt32(pointer, 72, fields);
        }
        finally { GlobalUnlock(handle); }
    }
    internal async Task PrintAsync(PrintSession session, CancellationToken token)
    {
        ObjectDisposedException.ThrowIf(_dc == 0, this);
        await using var lease = session.RetainPdf();
        await Task.Run(() => PrintCore(session, token), token);
    }
    private void PrintCore(PrintSession session, CancellationToken token)
    {
        var started = false;
        AbortCallback abort = (_, _) => !token.IsCancellationRequested;
        try
        {
            token.ThrowIfCancellationRequested(); Require(SetAbortProc(_dc, abort), "无法设置打印取消回调。");
            var info = new DocInfo { Size = Marshal.SizeOf<DocInfo>(), Name = session.Job.FileName };
            Require(StartDocW(_dc, ref info), "无法开始打印任务。"); started = true;
            foreach (var page in _pages)
            {
                token.ThrowIfCancellationRequested();
                using var bitmap = session.RenderPageAsync(page, _layout.RasterDpi, token).GetAwaiter().GetResult();
                Require(StartPage(_dc), "无法开始打印页面。");
                var bitmapInfo = new BitmapInfo { Size = (uint)Marshal.SizeOf<BitmapInfo>(), Width = bitmap.Width, Height = -bitmap.Height, Planes = 1, BitCount = 32 };
                // Skia uses tightly packed opaque BGRA; negative height is the GDI top-down DIB contract.
                if (bitmap.RowBytes != bitmap.Width * 4) throw new InvalidDataException("打印像素步幅无效。");
                var lines = StretchDIBits(_dc, _layout.X, _layout.Y, _layout.Width, _layout.Height, 0, 0, bitmap.Width, bitmap.Height, bitmap.GetPixels(), ref bitmapInfo, 0, 0x00CC0020);
                if (lines is 0 or -1) throw new Win32Exception(Marshal.GetLastWin32Error(), "无法传输打印页面。");
                token.ThrowIfCancellationRequested(); Require(EndPage(_dc), "打印页面提交失败。");
            }
            token.ThrowIfCancellationRequested(); Require(EndDoc(_dc), "打印任务提交失败。"); started = false;
        }
        finally { if (started) AbortDoc(_dc); GC.KeepAlive(abort); }
    }
    private static void Require(int result, string message) { if (result <= 0) throw new Win32Exception(Marshal.GetLastWin32Error(), message); }
    public void Dispose() { var dc = Interlocked.Exchange(ref _dc, 0); if (dc != 0) DeleteDC(dc); }
    private const uint ReturnDc = 0x100, ReturnDefault = 0x400, UseDriverCopies = 0x40000, NoSelection = 4, PageNumbers = 2, NoCurrentPage = 0x800000, HidePrintToFile = 0x100000, DisablePrintToFile = 0x80000;
    [StructLayout(LayoutKind.Sequential)] private struct PageRange { public uint From, To; }
    [StructLayout(LayoutKind.Sequential)] private struct PrintDialog
    {
        public uint Size; public nint Owner, DevMode, DevNames, Dc; public uint Flags, Flags2, ExclusionFlags, RangeCount, MaxRanges; public nint Ranges;
        public uint MinPage, MaxPage, Copies; public nint Instance, TemplateName, Callback; public uint PropertyPageCount; public nint PropertyPages; public uint StartPage, ResultAction;
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct DocInfo { public int Size; [MarshalAs(UnmanagedType.LPWStr)] public string Name; public nint Output, DataType; public uint Type; }
    [StructLayout(LayoutKind.Sequential)] private struct BitmapInfo { public uint Size; public int Width, Height; public ushort Planes, BitCount; public uint Compression, SizeImage; public int XPelsPerMeter, YPelsPerMeter; public uint ColorsUsed, ColorsImportant; }
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] [return: MarshalAs(UnmanagedType.Bool)] private delegate bool AbortCallback(nint dc, int error);
    [DllImport("comdlg32.dll", ExactSpelling = true)] private static extern int PrintDlgExW(ref PrintDialog dialog);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern nint GlobalLock(nint handle);
    [DllImport("kernel32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GlobalUnlock(nint handle);
    [DllImport("kernel32.dll")] private static extern nint GlobalFree(nint handle);
    [DllImport("kernel32.dll")] private static extern nuint GlobalSize(nint handle);
    [DllImport("gdi32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool DeleteDC(nint dc);
    [DllImport("gdi32.dll")] private static extern int GetDeviceCaps(nint dc, int index);
    [DllImport("gdi32.dll", SetLastError = true)] private static extern int SetAbortProc(nint dc, AbortCallback callback);
    [DllImport("gdi32.dll", ExactSpelling = true, SetLastError = true)] private static extern int StartDocW(nint dc, ref DocInfo info);
    [DllImport("gdi32.dll", SetLastError = true)] private static extern int StartPage(nint dc);
    [DllImport("gdi32.dll", SetLastError = true)] private static extern int EndPage(nint dc);
    [DllImport("gdi32.dll", SetLastError = true)] private static extern int EndDoc(nint dc);
    [DllImport("gdi32.dll")] private static extern int AbortDoc(nint dc);
    [DllImport("gdi32.dll", SetLastError = true)] private static extern int StretchDIBits(nint dc, int x, int y, int width, int height, int sourceX, int sourceY, int sourceWidth, int sourceHeight, nint pixels, ref BitmapInfo info, uint usage, uint operation);
}
