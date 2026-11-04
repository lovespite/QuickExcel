[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardNative.cs')
$root=Split-Path $PSScriptRoot -Parent
$out=Join-Path $root ('artifacts/clipboard/p0-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $out)
$excel=$null;$book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false
    $cases=@()
    foreach($date1904 in @($false,$true)) {
        $book=$excel.Workbooks.Add();$book.Date1904=$date1904;$sheet=$book.Worksheets.Item(1)
        $range=$sheet.Range('C4:F9');$range.Font.Name='Arial';$range.Font.Size=12
        $sheet.Range('C4').NumberFormat='@';$sheet.Range('C4').Value2='00123'
        $sheet.Range('D4').Value2=[double]0.12;$sheet.Range('D4').NumberFormat='0.00%'
        $sheet.Range('E4').Value2=$true;$sheet.Range('F4').Formula='=1/0'
        $sheet.Range('C5').Value2="中文`n行`t列";$sheet.Range('C5').WrapText=$true
        $sheet.Range('D5').Value2=[double]45000.5;$sheet.Range('D5').NumberFormat='yyyy-mm-dd hh:mm:ss'
        $sheet.Range('E5').Formula='=D4+$D$4+$D4+D$4'
        $sheet.Range('C6:D7').Merge();$sheet.Range('C6').Value2='merged'
        $sheet.Range('C4').Font.Bold=$true;$sheet.Range('C4').Font.Color=255;$sheet.Range('C4').Interior.Color=65535
        $sheet.Range('D4').Borders.LineStyle=1
        $excel.CalculateFullRebuild()
        $book.SaveAs((Join-Path $out "source-$date1904.xlsx"),51)
        $range.Copy()
        $payloads=[ClipboardNative]::Read();$formats=@()
        foreach($name in $payloads.Keys) {
            $bytes=$payloads[$name];$formats+=@{Name=$name;Bytes=$bytes.Length}
            if($bytes.Length) {[IO.File]::WriteAllBytes((Join-Path $out ("$date1904-"+($name -replace '[^a-zA-Z0-9]','_')+'.bin')),$bytes)}
        }
        $cases+=@{Date1904=$date1904;Formats=$formats}
        $excel.CutCopyMode=0;$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
    }
    @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Cases=$cases}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $out 'formats.json') -Encoding utf8
    Write-Output $out
} finally {if($book){$book.Close($false)};if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
