using System.Text;

namespace QuickExcel.Desktop.Avalonia.Diagnostics;

/// <summary>Independent PDF syntax fixture: no Skia-generated ToUnicode mapping or font dependency.</summary>
internal static class PdfTextFixture
{
    internal const string Expected = "plain文⽂ffiouterafterend";
    internal static void Write(string path)
    {
        const string content = """
            BT /F1 12 Tf 10 100 Td (plain) Tj
            /Span << /ActualText <FEFF6587> >> BDC (a) Tj EMC
            /Span << /ActualText <FEFF2F42> >> BDC (a) Tj EMC
            /Span << /ActualText (ffi) >> BDC (f) Tj EMC
            /Span << /ActualText (outer) >> BDC (x) Tj
            /Span << /ActualText (inner) >> BDC (y) Tj EMC (z) Tj EMC
            /Span BMC (after) Tj EMC
            /Span << /ActualText () >> BDC (hidden) Tj EMC
            (end) Tj ET
            """;
        string[] objects = ["<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", $"<< /Length {Encoding.ASCII.GetByteCount(content)} >>\nstream\n{content}\nendstream"];
        var pdf = new StringBuilder("%PDF-1.7\n"); var offsets = new List<int>();
        for (var i = 0; i < objects.Length; i++) { offsets.Add(pdf.Length); pdf.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n"); }
        var xref = pdf.Length; pdf.Append($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) pdf.Append($"{offset:D10} 00000 n \n");
        pdf.Append($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        File.WriteAllText(path, pdf.ToString(), Encoding.ASCII);
    }
}
