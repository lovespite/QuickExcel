[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run = Join-Path $root ('artifacts/clipboard/filter-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $run)
function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) }
}
$excel = $null; $cases = @()
Push-Location $root
try {
    if (-not $SkipBuild) {
        dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -v minimal
        if ($LASTEXITCODE -ne 0) { throw 'Clipboard filter build failed.' }
    }
    $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
    foreach ($scenario in 'filter', 'filter-header-only', 'filtered', 'table', 'table-filtered', 'filter-note') {
        $book = $null; $sheet = $null; $range = $null; $copy = $null; $process = $null
        $dir = Join-Path $run $scenario; [void](New-Item -ItemType Directory -Path $dir)
        try {
            $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
            $headers = @('编号','类别','金额','日期')
            for ($col=0; $col -lt 4; $col++) { $sheet.Cells.Item(4,3+$col).Value2 = $headers[$col] }
            for ($row=5; $row -le 9; $row++) {
                $sheet.Cells.Item($row,3).NumberFormat = '@'; $sheet.Cells.Item($row,3).Value2 = '00' + $row
                $sheet.Cells.Item($row,4).Value2 = $(if ($row % 2) {'保留'} else {'隐藏'})
                $sheet.Cells.Item($row,5).Value2 = [double]($row / 100); $sheet.Cells.Item($row,5).NumberFormat = '0.00%'
                $sheet.Cells.Item($row,6).Value2 = [double](45000.5+$row); $sheet.Cells.Item($row,6).NumberFormat = 'yyyy-mm-dd hh:mm:ss'
            }
            $range = $sheet.Range('C4:F9')
            if ($scenario.StartsWith('table')) { $table = $sheet.ListObjects.Add(1,$range,[Type]::Missing,1); $table.TableStyle = ''; Release-Com $table }
            else { [void]$range.AutoFilter() }
            $header = $sheet.Range('C4:F4'); $font = $header.Font; $font.Bold = $true; Release-Com $font; Release-Com $header
            if ($scenario.EndsWith('filtered')) { [void]$range.AutoFilter(2,'保留') }
            if ($scenario -eq 'filter-note') { $cell = $sheet.Range('C5'); $note = $cell.AddComment('说明'); Release-Com $note; Release-Com $cell }
            $book.SaveAs((Join-Path $dir 'source.xlsx'),51)
            $expected = @(); $outputRow = 3; $lastRow = if ($scenario -eq 'filter-header-only') {4} else {9}
            for ($row=4; $row -le $lastRow; $row++) {
                $sourceRow = $sheet.Rows.Item($row)
                try { $hidden = [bool]$sourceRow.Hidden } finally { Release-Com $sourceRow }
                if ($hidden) { continue }
                for ($col=3; $col -le 6; $col++) {
                    $cell = $sheet.Cells.Item($row,$col); $font = $cell.Font
                    try {
                        $value = $cell.Value2; $kind = if ($value -is [string]) {'Text'} else {'Number'}
                        $text = if ($kind -eq 'Number') { ([double]$value).ToString('R',[Globalization.CultureInfo]::InvariantCulture) } else { [string]$value }
                        $expected += @{row=$outputRow; column=$col-1; kind=$kind; value=$text; bold=[bool]$font.Bold}
                    } finally { Release-Com $font; Release-Com $cell }
                }
                $outputRow++
            }
            $expected | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $dir 'expected.json') -Encoding utf8
            $request = @{operation='import'; mode='all'}
            if ($scenario -eq 'filter') { $request.expectedPasteNotice = '筛选' }
            if ($scenario -eq 'filter-note') { $request.expectedPasteNotice = '自动改为只贴值' }
            $request | ConvertTo-Json | Set-Content (Join-Path $dir 'request.json') -Encoding utf8
            $copy = $sheet.Range('C4:F' + $lastRow); [void]$copy.Copy()
            $exe = Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/QuickExcel.Avalonia.exe'
            $process = Start-Process -FilePath $exe -ArgumentList @('--clipboard-exchange', ('"{0}"' -f $dir)) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $dir 'stdout.log') -RedirectStandardError (Join-Path $dir 'stderr.log')
            if (-not $process.WaitForExit(60000)) { throw 'Clipboard exchange exceeded 60 seconds.' }
            $result = Get-Content (Join-Path $dir 'result.json') -Raw | ConvertFrom-Json
            if ($process.ExitCode -ne 0 -or -not $result.passed) { throw $result.error }
            $actual = @(Get-Content (Join-Path $dir 'model.json') -Raw | ConvertFrom-Json)
            $differences = @()
            if ($actual.Count -ne $expected.Count) { $differences += @{Property='CellCount'; Expected=$expected.Count; Actual=$actual.Count} }
            foreach ($cell in $expected) {
                $found = @($actual | Where-Object { $_.row -eq $cell.row -and $_.column -eq $cell.column })
                if ($found.Count -ne 1) { $differences += @{Property='Position'; Expected=$cell}; continue }
                $a = $found[0]; $bold = if ($scenario -eq 'filter-note') {$false} else {$cell.bold}
                if ($a.kind -cne $cell.kind -or $a.value -cne $cell.value -or $a.style.Bold -ne $bold -or $null -ne $a.formula) {
                    $differences += @{Property='TypedValueOrStyle'; Expected=$cell; ExpectedBold=$bold; Actual=$a}
                }
            }
            $cases += @{Name=$scenario; Passed=($differences.Count -eq 0); Differences=$differences; Path=$dir}
            Write-Output "$scenario`: $($differences.Count) differences"
        } catch {
            $cases += @{Name=$scenario; Passed=$false; Error=$_.Exception.ToString(); Path=$dir}
            Write-Output "$scenario`: FAILED; $($_.Exception.Message)"
        } finally {
            if ($process) { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
            Release-Com $copy; Release-Com $range; Release-Com $sheet
            if ($book) { $book.Close($false); Release-Com $book }
        }
    }
    @{Version=$excel.Version; Build=$excel.Build; Language=$excel.LanguageSettings.LanguageID(2); Passed=(@($cases | Where-Object {-not $_.Passed}).Count -eq 0); Cases=$cases;
        Scope='Synthetic Excel workbooks, real OS clipboard and Avalonia paste handlers; typed values, visible row mapping, styles, fallback notice, undo/redo and save/reopen'} |
        ConvertTo-Json -Depth 12 | Set-Content (Join-Path $run 'report.json') -Encoding utf8
    Write-Output "Evidence: $run"
    if (@($cases | Where-Object {-not $_.Passed}).Count) { throw 'Filter clipboard cases failed; see report.json.' }
} finally {
    if ($excel) { $excel.Quit(); Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Pop-Location
}
