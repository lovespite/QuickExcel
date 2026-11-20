[CmdletBinding()]
param([switch]$SkipBuild, [switch]$ProbeOnly, [ValidateSet('Xml','Html')][string]$ImportFormat = 'Xml')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run=Join-Path $root ('artifacts/clipboard/styles-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $run)
Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardNative.cs')
function Release-Com($value) {if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
$specs=@(
    @{Name='single';Underline=2},@{Name='double';Underline=-4119},@{Name='single-accounting';Underline=4},@{Name='double-accounting';Underline=5},
    @{Name='strike';Strike=$true},@{Name='both';Strike=$true;Underline=-4119},
    @{Name='hairline';Weight=1},@{Name='thin';Weight=2},@{Name='medium';Weight=-4138},@{Name='thick';Weight=4},
    @{Name='dashed';Line=-4115;Weight=2},@{Name='dotted';Line=-4118;Weight=2},@{Name='double-border';Line=-4119;Weight=4},
    @{Name='indent-1';Indent=1},@{Name='indent-3';Indent=3},@{Name='indent-15';Indent=15},
    @{Name='rotation-negative';Rotation=-45},@{Name='rotation-positive';Rotation=45},@{Name='rotation-minus90';Rotation=-90},@{Name='rotation-90';Rotation=90},
    @{Name='stacked';Rotation=-4166},@{Name='combined';Rotation=30;Indent=2;Underline=2;Strike=$true;Weight=-4138},@{Name='plain'})
function Capture($sheet,[int]$row,[int]$col) {
    $result=@()
    for($i=0;$i -lt $specs.Count;$i++) {
        $cell=$sheet.Cells.Item($row+$i,$col);$font=$cell.Font
        try {
            $borders=@();foreach($edge in 7,8,9,10){$b=$cell.Borders.Item($edge);try{$borders+=@{Line=$b.LineStyle;Weight=$b.Weight;Color=$b.Color}}finally{Release-Com $b}}
            $result+=@{Name=$specs[$i].Name;Value=[string]$cell.Value2;Underline=$font.Underline;Strike=$font.Strikethrough;Indent=$cell.IndentLevel;Rotation=$cell.Orientation;Borders=$borders}
        } finally {Release-Com $font;Release-Com $cell}
    }
    return ,$result
}
$excel=$null;$book=$null;$targetBook=$null;$sheet=$null;$target=$null;$process=$null
Push-Location $root
try {
    if(-not $SkipBuild -and -not $ProbeOnly){dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -v minimal;if($LASTEXITCODE -ne 0){throw 'Build failed'}}
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
    for($i=0;$i -lt $specs.Count;$i++) {
        $s=$specs[$i];$cell=$sheet.Cells.Item(4+$i,3);$font=$cell.Font
        try {
            $cell.Value2=$s.Name;$cell.HorizontalAlignment=-4131
            if($s.ContainsKey('Underline')){$font.Underline=$s.Underline}
            if($s.ContainsKey('Strike')){$font.Strikethrough=$s.Strike}
            if($s.ContainsKey('Indent')){$cell.IndentLevel=$s.Indent}
            if($s.ContainsKey('Rotation')){$cell.Orientation=$s.Rotation}
            if($s.ContainsKey('Weight')) {foreach($edge in 7,8,9,10){$b=$cell.Borders.Item($edge);try{$b.LineStyle=$(if($s.ContainsKey('Line')){$s.Line}else{1});$b.Weight=$s.Weight;$b.Color=255}finally{Release-Com $b}}}
        } finally {Release-Com $font;Release-Com $cell}
    }
    $book.SaveAs((Join-Path $run 'source.xlsx'),51);$expected=Capture $sheet 4 3
    $expected|ConvertTo-Json -Depth 10|Set-Content (Join-Path $run 'expected.json') -Encoding utf8
    $range=$sheet.Range('C4:C'+(3+$specs.Count));[void]$range.Copy();Release-Com $range
    $payload=[ClipboardNative]::Read();foreach($name in $payload.Keys){if($payload[$name].Length){[IO.File]::WriteAllBytes((Join-Path $run (($name -replace '[^a-zA-Z0-9]','_')+'.bin')),$payload[$name])}}
    if($ImportFormat -eq 'Html'){$formats=[byte[][]]::new(1);$formats[0]=$payload['HTML Format'];[ClipboardNative]::Write(@('HTML Format'),$formats)}
    if(-not $ProbeOnly) {
        @{operation='import';mode=$(if($ImportFormat -eq 'Html'){'formats'}else{'all'})}|ConvertTo-Json|Set-Content (Join-Path $run 'request.json') -Encoding utf8
        $exe=Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/quickxl.exe'
        $process=Start-Process -FilePath $exe -ArgumentList @('--clipboard-exchange',('"{0}"' -f $run)) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $run 'stdout.log') -RedirectStandardError (Join-Path $run 'stderr.log')
        if(-not $process.WaitForExit(60000)){throw 'Style exchange timed out'}
        $result=Get-Content (Join-Path $run 'result.json') -Raw|ConvertFrom-Json
        if($process.ExitCode -ne 0 -or -not $result.passed){throw $result.error}
        $targetBook=$excel.Workbooks.Add();$target=$targetBook.Worksheets.Item(1);[void]$target.Paste($target.Range('G10'));$actual=Capture $target 10 7
        $differences=@();for($i=0;$i -lt $specs.Count;$i++){foreach($key in 'Value','Underline','Strike','Indent','Rotation','Borders'){
            if($ImportFormat -eq 'Html' -and $key -eq 'Value'){continue}
            if(($expected[$i][$key]|ConvertTo-Json -Depth 5 -Compress) -cne ($actual[$i][$key]|ConvertTo-Json -Depth 5 -Compress)){$differences+=@{Name=$specs[$i].Name;Property=$key;Expected=$expected[$i][$key];Actual=$actual[$i][$key]}}
        }}
        $targetBook.SaveAs((Join-Path $run 'excel-result.xlsx'),51)
        @{Passed=($differences.Count -eq 0);ImportFormat=$ImportFormat;Version=$excel.Version;Build=$excel.Build;Cases=$specs.Count;Differences=$differences;Expected=$expected;Actual=$actual}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $run 'report.json') -Encoding utf8
        if($differences.Count){throw "Styles: $($differences.Count) differences. Evidence: $run"}
        Write-Output "PASS: $($specs.Count) Excel style cases"
    }
    Write-Output "Evidence: $run"
} finally {
    if($process){if(-not $process.HasExited){$process.Kill()};$process.Dispose()}
    Release-Com $target;Release-Com $sheet
    if($targetBook){$targetBook.Close($false);Release-Com $targetBook};if($book){$book.Close($false);Release-Com $book}
    if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();Pop-Location
}
