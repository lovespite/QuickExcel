[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run=Join-Path $root ('artifacts/clipboard/object-exchange-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $run)
Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardNative.cs')
Add-Type -AssemblyName System.Drawing
function Release-Com($value) {if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
function Exchange([string]$name,[string]$mode) {
    $dir=Join-Path $run $name;[void](New-Item -ItemType Directory -Path $dir)
    @{operation='object';mode=$mode}|ConvertTo-Json|Set-Content (Join-Path $dir 'request.json') -Encoding utf8
    $exe=Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/QuickExcel.Avalonia.exe'
    $process=Start-Process -FilePath $exe -ArgumentList @('--clipboard-exchange',('"{0}"' -f $dir)) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $dir 'stdout.log') -RedirectStandardError (Join-Path $dir 'stderr.log')
    try {
        if(-not $process.WaitForExit(60000)){throw 'Object exchange timed out'}
        $result=Get-Content (Join-Path $dir 'result.json') -Raw|ConvertFrom-Json
        if($process.ExitCode -ne 0 -or -not $result.passed){throw $result.error}
    } finally {if(-not $process.HasExited){$process.Kill()};$process.Dispose()}
    return $dir
}
$excel=$null;$book=$null;$sheet=$null;$chartObject=$null;$chart=$null;$picture=$null;$targetBook=$null;$target=$null
$results=@()
Push-Location $root
try {
    if(-not $SkipBuild){dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -v minimal;if($LASTEXITCODE -ne 0){throw 'Build failed'}}
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
    $matrix=[object[,]]::new(4,3);$matrix[0,0]='Category';$matrix[0,1]='Revenue';$matrix[0,2]='Cost'
    for($r=1;$r -lt 4;$r++){$matrix[$r,0]='Q'+$r;$matrix[$r,1]=[double](10*$r);$matrix[$r,2]=[double](7*$r)}
    $range=$sheet.Range('A1:C4');$range.Value2=$matrix
    $objects=$sheet.ChartObjects();$chartObject=$objects.Add(100,100,480,300);Release-Com $objects
    $chart=$chartObject.Chart;$chart.ChartType=51;$chart.SetSourceData($range);$chart.HasTitle=$true;$chart.ChartTitle.Text='Clipboard Chart';Release-Com $range
    $book.SaveAs((Join-Path $run 'source.xlsx'),51)
    foreach($case in @(@{Name='column';Type=51;Mode='chart'},@{Name='line';Type=4;Mode='chart'},@{Name='pie';Type=5;Mode='chart'},@{Name='scatter';Type=-4169;Mode='chart'},@{Name='chart-image';Type=51;Mode='image'},
        @{Name='area-fallback';Type=1;Mode='all';Fallback=$true},@{Name='combo-fallback';Type=51;Mode='chart';Fallback=$true})) {
        $book.Activate();for($r=1;$r -lt 4;$r++){$cell=$sheet.Cells.Item($r+1,1);$cell.Value2=$(if($case.Type -eq -4169){[double]$r}else{'Q'+$r});Release-Com $cell}
        $chart.ChartType=$case.Type;$range=$sheet.Range($(if($case.Type -eq 5){'A1:B4'}else{'A1:C4'}));$chart.SetSourceData($range);Release-Com $range
        if($case.Name -eq 'combo-fallback'){$series=$chart.SeriesCollection(2);try{$series.ChartType=4}finally{Release-Com $series}}
        $chartObject.Copy();$dir=Exchange $case.Name $case.Mode
        if($case.ContainsKey('Fallback') -and (Get-Content (Join-Path $dir 'paste-status.txt') -Raw) -notmatch '自动降级为图片'){throw 'Unsupported chart did not report image fallback'}
        if($case.Name -eq 'column'){[void](Exchange 'column-private' 'chart');$results+=@{Name='column-private-cross-instance';Passed=$true}}
        $targetBook=$excel.Workbooks.Add();$target=$targetBook.Worksheets.Item(1);$target.Activate();$target.Paste()
        $charts=$target.ChartObjects();$count=$charts.Count;Release-Com $charts
        if($case.Mode -eq 'chart' -and -not $case.ContainsKey('Fallback')) {
            if($count -ne 1){throw "Expected editable chart in Excel, got $count"}
            $copy=$target.ChartObjects(1);$copyChart=$copy.Chart;$series=$copyChart.SeriesCollection(1)
            try{
                if(($series.Values -join ',') -ne '10,20,30'){throw 'Excel chart values mismatch'}
                if($series.Formula -match '!|\['){throw 'External chart still depends on source or destination cells'}
                $copyChart.ChartTitle.Text='Edited after paste';$series.Values=[double[]]@(11,22,33)
                if(($series.Values -join ',') -ne '11,22,33'){throw 'Excel chart is not editable'}
            }finally{Release-Com $series;Release-Com $copyChart;Release-Com $copy}
            $data=Get-Content (Join-Path $dir 'chart-data.json') -Raw|ConvertFrom-Json
            $categories=if($case.Type -eq -4169){'1,2,3'}else{'Q1,Q2,Q3'}
            if(($data[0].Values -join ',') -ne '10,20,30' -or ($data[0].Categories -join ',') -ne $categories){throw 'Imported editable chart data mismatch'}
        } elseif($count -ne 0 -or $target.Shapes.Count -ne 1){throw 'Expected one picture, no editable chart'}
        $targetBook.SaveAs((Join-Path $dir 'excel-result.xlsx'),51)
        $results+=@{Name=$case.Name;Passed=$true};Release-Com $target;$target=$null;$targetBook.Close($false);Release-Com $targetBook;$targetBook=$null
    }
    $bitmap=[Drawing.Bitmap]::new(80,40);$graphics=[Drawing.Graphics]::FromImage($bitmap)
    try{$graphics.Clear([Drawing.Color]::CornflowerBlue);$bitmap.Save((Join-Path $run 'fixture.png'),[Drawing.Imaging.ImageFormat]::Png)}finally{$graphics.Dispose();$bitmap.Dispose()}
    $book.Activate();$picture=$sheet.Shapes.AddPicture((Join-Path $run 'fixture.png'),0,-1,10,10,80,40);$picture.Copy();$dir=Exchange 'image' 'all'
    $targetBook=$excel.Workbooks.Add();$target=$targetBook.Worksheets.Item(1);$target.Paste()
    if($target.Shapes.Count -ne 1){throw 'Excel did not receive an image'}
    $targetBook.SaveAs((Join-Path $dir 'excel-result.xlsx'),51);$results+=@{Name='image';Passed=$true}
    @{Passed=$true;Version=$excel.Version;Build=$excel.Build;Cases=$results}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $run 'report.json') -Encoding utf8
    Write-Output "PASS: $($results.Count) Excel object exchanges. Evidence: $run"
} finally {
    Release-Com $target;if($targetBook){$targetBook.Close($false);Release-Com $targetBook}
    Release-Com $picture;Release-Com $chart;Release-Com $chartObject;Release-Com $sheet
    if($book){$book.Close($false);Release-Com $book};if($excel){$excel.Quit();Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();Pop-Location
}
