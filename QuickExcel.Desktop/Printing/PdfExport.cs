using System.IO;
using System.Runtime.InteropServices;
using System.Printing;
using System.Windows.Xps.Packaging;

namespace QuickExcel.Desktop;

/// <summary>Exports the exact preview paginator through the installed Windows PDF driver.</summary>
internal static class PdfExport
{
    internal static async Task WriteAsync(WorkbookPaginator paginator, string destination, CancellationToken cancellationToken = default)
    {
        var fullPath = Path.GetFullPath(destination); var suffix = Guid.NewGuid().ToString("N");
        var xpsPath = fullPath + "." + suffix + ".xps"; var pdfPath = fullPath + "." + suffix + ".pdf";
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var settings = new PrintTicket { PageMediaSize = new PageMediaSize(paginator.PageSize.Width, paginator.PageSize.Height), PageOrientation = PageOrientation.Portrait };
            using (var document = new XpsDocument(xpsPath, FileAccess.ReadWrite))
                XpsDocument.CreateXpsDocumentWriter(document).Write(paginator, settings);
            using var ticketXml = settings.GetXmlStream(); var ticketBytes = ticketXml.ToArray();
            if (new FileInfo(xpsPath).Length > 64 * 1024 * 1024) throw new InvalidOperationException("打印任务超过 64 MiB，请缩小打印范围。");
            await Task.Run(() => Print(xpsPath, pdfPath, ticketBytes, cancellationToken), cancellationToken);
            using (var input = File.OpenRead(pdfPath))
            {
                var header = new byte[5]; if (input.Read(header) != 5 || !header.SequenceEqual("%PDF-"u8.ToArray())) throw new IOException("PDF 驱动未生成有效文件。");
            }
            cancellationToken.ThrowIfCancellationRequested(); File.Move(pdfPath, fullPath, true);
        }
        finally
        {
            if (File.Exists(xpsPath)) File.Delete(xpsPath);
            if (File.Exists(pdfPath)) File.Delete(pdfPath);
        }
    }
    private static void Print(string xpsPath, string pdfPath, byte[] ticketBytes, CancellationToken cancellationToken)
    {
        using var complete = new EventWaitHandle(false, EventResetMode.ManualReset);
        IXpsPrintJob? job = null; IXpsPrintJobStream? document = null; IXpsPrintJobStream? ticket = null;
        try
        {
            Marshal.ThrowExceptionForHR(StartXpsPrintJob("Microsoft Print to PDF", "QuickExcel PDF", pdfPath, IntPtr.Zero, complete.SafeWaitHandle.DangerousGetHandle(), IntPtr.Zero, 0, out job, out document, out ticket));
            ticket.Write(ticketBytes, (uint)ticketBytes.Length, out var ticketWritten);
            if (ticketWritten != ticketBytes.Length) throw new IOException("打印设置未完整写入。");
            ticket.Close();
            using (var input = File.OpenRead(xpsPath))
            {
                var buffer = new byte[65536]; int count;
                while ((count = input.Read(buffer)) > 0)
                {
                    cancellationToken.ThrowIfCancellationRequested(); document.Write(buffer, (uint)count, out var written);
                    if (written != count) throw new IOException("打印流未完整写入。");
                }
            }
            document.Close(); var started = Environment.TickCount64;
            while (!complete.WaitOne(100))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (Environment.TickCount64 - started > 60000) throw new TimeoutException("Windows PDF 驱动超时，未替换目标文件。");
            }
            job.GetJobStatus(out var status);
            if (status.Completion != 1 || status.Error < 0) throw new IOException($"Windows PDF 打印失败（0x{status.Error:X8}）。");
        }
        catch
        {
            if (job is not null) { try { job.Cancel(); complete.WaitOne(5000); } catch (Exception ex) when (ex is COMException or ArgumentException) { } }
            throw;
        }
        finally
        {
            if (ticket is not null) Marshal.FinalReleaseComObject(ticket);
            if (document is not null) Marshal.FinalReleaseComObject(document);
            if (job is not null) Marshal.FinalReleaseComObject(job);
        }
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct JobStatus { public uint JobId, CurrentDocument, CurrentPage, CurrentPageTotal; public int Completion, Error; }
    [ComImport, Guid("7a77dc5f-45d6-4dff-9307-d8cb846347ca"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IXpsPrintJobStream
    {
        void Read([Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] buffer, uint count, out uint read);
        void Write([In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] buffer, uint count, out uint written);
        void Close();
    }
    [ComImport, Guid("5ab89b06-8194-425f-ab3b-d7a96e350161"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IXpsPrintJob { void Cancel(); void GetJobStatus(out JobStatus status); }
    [DllImport("XpsPrint.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern int StartXpsPrintJob(string printerName, string jobName, string outputFileName, IntPtr progressEvent, IntPtr completionEvent,
        IntPtr printablePages, uint printablePageCount, out IXpsPrintJob job, out IXpsPrintJobStream documentStream, out IXpsPrintJobStream ticketStream);
}
