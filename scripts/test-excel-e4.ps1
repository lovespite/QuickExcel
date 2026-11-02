param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$artifactRoot = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($artifactRoot, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $target)
$checks = [Collections.Generic.List[string]]::new()
function Check($ok, [string]$name) { if (-not $ok) { throw $name }; $checks.Add($name) }
function Release-Com($object) { if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) } }
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add()
    while ($book.Worksheets.Count -gt 1) { $extra = $book.Worksheets.Item($book.Worksheets.Count); $extra.Delete(); Release-Com $extra }
    $sheet = $book.Worksheets.Item(1); $sheet.Name = 'Report'
    $sheet.Cells.Item(1,1) = 'Month'; $sheet.Cells.Item(1,2) = 'Revenue'; $sheet.Cells.Item(1,3) = 'Cost'
    for ($i=2; $i -le 5; $i++) { $sheet.Cells.Item($i,1) = "M$($i-1)"; $sheet.Cells.Item($i,2) = $i * 12.5; $sheet.Cells.Item($i,3) = $i * 8 }
    $offset = 0
    foreach ($kind in @(51,4,5)) {
        $object = $sheet.ChartObjects().Add(20,150+$offset,420,220); $chart = $object.Chart
        $chart.ChartType = $kind; $chart.SetSourceData($sheet.Range($(if ($kind -eq 5) {'A1:B5'} else {'A1:C5'})))
        $chart.HasTitle = $true; $chart.ChartTitle.Text = "Report $kind"; $offset += 240
        Release-Com $chart; Release-Com $object
    }
    $sheet.Range('A1:C5').Borders.LineStyle = 1; $sheet.Range('A1:C1').Font.Bold = $true
    $sheet.PageSetup.PrintArea = '$A$1:$I$62'; $sheet.PageSetup.Orientation = 2; $sheet.PageSetup.PaperSize = 9
    $sheet.PageSetup.Zoom = $false; $sheet.PageSetup.FitToPagesWide = 1; $sheet.PageSetup.FitToPagesTall = 2
    $sheet.PageSetup.PrintTitleRows = '$1:$1'; $sheet.PageSetup.CenterHeader = 'E4 report'; $sheet.PageSetup.RightFooter = '&P / &N'
    $sheet.ExportAsFixedFormat(0, (Join-Path $target 'excel-source-report.pdf'))
    $pictures = $book.Worksheets.Add([Type]::Missing, $sheet); $pictures.Name = 'Pictures'
    $png = Join-Path $target 'fixture.png'
    [IO.File]::WriteAllBytes($png, [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/mWQAAAAASUVORK5CYII='))
    for ($i=0; $i -lt 3; $i++) {
        $picture = $pictures.Shapes.AddPicture($png, 0, -1, (37.5 + $i * 150), 30.75, 121.5, 60.25)
        $picture.Placement = $i + 1; Release-Com $picture
    }
    $protected = $book.Worksheets.Add([Type]::Missing, $pictures); $protected.Name = 'Protected'
    $protected.Range('A1') = 'Locked'; $protected.Range('B2') = 5; $protected.Range('B2').Locked = $false
    $protected.Range('A2').Formula = '=1+2'; $protected.Range('A2').FormulaHidden = $true
    $note = $protected.Range('A1').AddComment('Original note'); Release-Com $note
    $note = $protected.Range('D4').AddComment('Remove me'); Release-Com $note
    # Passwordless protection: permit formatting rows, sorting, existing filters and inserting rows.
    $protected.Protect('', $true, $true, $true, $false, $false, $false, $true, $false, $true, $false, $false, $false, $true, $true, $false)
    $paged = $book.Worksheets.Add([Type]::Missing, $protected); $paged.Name = 'Paged'
    $paged.Range('A1') = 'Row'; $paged.Range('B1') = 'Value'
    for ($i=2; $i -le 36; $i++) { $paged.Cells.Item($i,1) = "Record $($i-1)"; $paged.Cells.Item($i,2) = $i * 10 }
    $paged.Range('C1:E2').Merge(); $paged.Range('C1') = 'Merged title'; $paged.Range('C1:E2').Interior.Color = 15773696
    $paged.Rows.RowHeight = 20; $paged.Range('A:E').ColumnWidth = 12
    $picture = $paged.Shapes.AddPicture($png, 0, -1, 170, 75, 100, 45); $picture.Placement = 2; Release-Com $picture
    $paged.PageSetup.PrintArea = '$A$1:$E$36'; $paged.PageSetup.PrintTitleRows = '$1:$1'; $paged.PageSetup.PaperSize = 9
    $paged.PageSetup.Zoom = 100; $paged.PageSetup.CenterHeader = 'Paged report'; $paged.PageSetup.RightFooter = '&P / &N'
    $pageBreak = $paged.HPageBreaks.Add($paged.Range('A21')); Release-Com $pageBreak
    $paged.ExportAsFixedFormat(0, (Join-Path $target 'excel-source-paged.pdf')); Release-Com $paged
    $sheet.Activate()
    $source = Join-Path $target 'excel-source.xlsx'; $book.SaveAs($source, 51)
    Release-Com $protected; Release-Com $pictures; Release-Com $sheet
    $book.Close($false); Release-Com $book; $book = $null
    $sourceHash = (Get-FileHash -LiteralPath $source).Hash
    & dotnet run --project (Join-Path $repo 'tools/QuickExcel.CompatibilityProbe/QuickExcel.CompatibilityProbe.csproj') -c Release --no-restore --artifacts-path (Join-Path $repo 'artifacts/build') -- e4-report $source (Join-Path $target 'quickexcel')
    if ($LASTEXITCODE -ne 0) { throw 'QuickExcel report probe failed.' }
    $expected = Get-Content -LiteralPath (Join-Path $target 'quickexcel/result.json') -Raw | ConvertFrom-Json
    $output = $expected.Outputs[0]
    $captures = @()
    for ($pass=0; $pass -lt 2; $pass++) {
        $path = if ($pass -eq 0) { $output.File } else { Join-Path $target 'excel-resaved.xlsx' }
        $book = $excel.Workbooks.Open($path, 0, $true)
        Check (-not $book.RepairMode) "Pass $pass normal Open: no repair mode"
        $sheet = $book.Worksheets.Item('Report'); $objects = $sheet.ChartObjects()
        Check ($objects.Count -eq 3) "Pass $pass three charts"
        Check ([double]$sheet.Range('B2').Value2 -eq 123.5) "Pass $pass edited source value"
        for ($i=1; $i -le 3; $i++) {
            $obj = $objects.Item($i); $chart = $obj.Chart; $series = $chart.SeriesCollection(1)
            Check ($chart.ChartType -eq @(51,4,5)[$i-1]) "Pass $pass chart $i type"
            Check ($chart.ChartTitle.Text -eq "Report $(@(51,4,5)[$i-1])") "Pass $pass chart $i title"
            Check ([double]@($series.Values)[0] -eq 123.5) "Pass $pass chart $i current values"
            Check ($series.Formula.Contains('Report!')) "Pass $pass chart $i references"
            $captures += [ordered]@{Pass=$pass; Chart=$i; Formula=$series.Formula; Values=@($series.Values); Left=$obj.Left; Top=$obj.Top; Width=$obj.Width; Height=$obj.Height}
            Release-Com $series; Release-Com $chart; Release-Com $obj
        }
        Check ($sheet.PageSetup.PrintTitleRows -eq '$1:$1') "Pass $pass repeat title"
        Check ($sheet.PageSetup.Orientation -eq 2 -and $sheet.PageSetup.FitToPagesTall -eq 2) "Pass $pass landscape and fit"
        if ($pass -eq 0) { $sheet.ExportAsFixedFormat(0, (Join-Path $target 'excel-output-report.pdf')) }
        Release-Com $objects; Release-Com $sheet
        $sheet = $book.Worksheets.Item('Pictures')
        Check ($sheet.Shapes.Count -eq 3) "Pass $pass picture count"
        for ($i=1; $i -le 3; $i++) {
            $shape = $sheet.Shapes.Item($i); $bounds = $output.Pictures[$i-1].Bounds
            $captures += [ordered]@{Pass=$pass; Picture=$i; Expected=$bounds; Left=$shape.Left; Top=$shape.Top; Width=$shape.Width; Height=$shape.Height; ColumnWidth=$sheet.StandardWidth; CellWidth=$sheet.Range('A1').Width}
            $captures | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $target 'captures.json')
            Check ($shape.Placement -eq $i) "Pass $pass picture $i behavior"
            foreach ($pair in @(@('Left','X'),@('Top','Y'),@('Width','Width'),@('Height','Height'))) {
                Check ([Math]::Abs([double]$shape.($pair[0]) - [double]$bounds.($pair[1]) * .75) -le 1) "Pass $pass picture $i $($pair[0]) within 1 pt"
            }
            Release-Com $shape
        }
        Release-Com $sheet; $sheet = $book.Worksheets.Item('Protected')
        Check ($sheet.ProtectContents -and $sheet.ProtectDrawingObjects) "Pass $pass protection retained"
        Check ($sheet.Protection.AllowInsertingRows -and $sheet.Protection.AllowFormattingRows -and $sheet.Protection.AllowSorting -and $sheet.Protection.AllowFiltering) "Pass $pass allowed operations retained"
        Check (-not $sheet.Range('B2').Locked -and $sheet.Range('B2').Value2 -eq 77 -and $sheet.Range('A2').FormulaHidden) "Pass $pass unlocked edit and hidden formula"
        Check ($sheet.Range('A1').Comment.Text() -eq 'Edited note') "Pass $pass edited note"
        Check ($sheet.Range('C3').Comment.Text() -eq 'New note' -and $sheet.Range('D4').Comment -eq $null) "Pass $pass inserted and removed notes"
        Release-Com $sheet
        $sheet = $book.Worksheets.Item('Paged')
        Check ($sheet.HPageBreaks.Item(1).Location.Row -eq 21) "Pass $pass manual page starts at row 21"
        Check ($sheet.PageSetup.PrintTitleRows -eq '$1:$1' -and $sheet.Range('C1:E2').MergeCells -and $sheet.Shapes.Count -eq 1) "Pass $pass paged titles, merge and image"
        if ($pass -eq 0) { $sheet.ExportAsFixedFormat(0, (Join-Path $target 'excel-output-paged.pdf')) }; Release-Com $sheet
        if ($pass -eq 0) { $book.SaveAs((Join-Path $target 'excel-resaved.xlsx'), 51) }
        $book.Close($false); Release-Com $book; $book = $null
    }
    Check ((Get-FileHash -LiteralPath $source).Hash -eq $sourceHash) 'Source fixture unchanged'
    [ordered]@{Status='PASS'; Version=$excel.Version; Build=$excel.Build; Language=$excel.LanguageSettings.LanguageID(2); Checks=$checks.Count; Results=$checks; Captures=$captures; SourceSha256=$sourceHash} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $target 'excel-result.json')
    "PASS: $($checks.Count) Excel E4 checks; $target"
} catch {
    $_ | Out-String | Set-Content -LiteralPath (Join-Path $target 'failure.txt'); throw
} finally {
    if ($null -ne $book) { $book.Close($false); Release-Com $book }
    if ($null -ne $excel) { $excel.Quit(); Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
