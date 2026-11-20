[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run=Join-Path $root ('artifacts/clipboard/interop-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $run)
Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardNative.cs')
function Release-Com($value) { if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)} }
function Run-QuickExcel([string]$name,[string]$operation,[string]$mode='all') {
    $dir=Join-Path $run $name;[void](New-Item -ItemType Directory -Path $dir)
    @{operation=$operation;mode=$mode}|ConvertTo-Json|Set-Content (Join-Path $dir 'request.json') -Encoding utf8
    $exe=Join-Path $root 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/quickxl.exe'
    $process=Start-Process -FilePath $exe -ArgumentList @('--clipboard-exchange',('"{0}"' -f $dir)) -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $dir 'stdout.log') -RedirectStandardError (Join-Path $dir 'stderr.log')
    try {
        if(-not $process.WaitForExit(60000)){throw "Clipboard exchange timed out: $dir"}
        $result=Get-Content (Join-Path $dir 'result.json') -Raw|ConvertFrom-Json
        if($process.ExitCode -ne 0 -or -not $result.passed){throw "QuickExcel failed: $($result.error) Evidence: $dir"}
    } finally {if(-not $process.HasExited){$process.Kill()};$process.Dispose()}
    return $dir
}
function Capture($sheet,[int]$top,[int]$left) {
    $result=@()
    for($r=0;$r -lt 6;$r++){for($c=0;$c -lt 4;$c++){
        $cell=$sheet.Cells.Item($top+$r,$left+$c);$font=$cell.Font;$fill=$cell.Interior;$merge=$cell.MergeArea
        try {
            $value=$cell.Value2;$kind=if($null -eq $value){'Blank'}elseif($value -is [bool]){'Boolean'}elseif($value -is [string]){'Text'}elseif($value -is [int] -and $value -lt -2000000000){'Error'}else{'Number'}
            if($kind -eq 'Error'){$value=$cell.Text}
            $formula=if($cell.HasFormula){[string]$cell.FormulaR1C1}else{$null}
            $border=@();foreach($edge in 7,8,9,10){$b=$cell.Borders.Item($edge);try{$border+=@{Style=$b.LineStyle;Weight=$b.Weight;Color=$b.Color}}finally{Release-Com $b}}
            $result+= [ordered]@{Row=$r;Column=$c;Kind=$kind;Value=$value;Text=[string]$cell.Text;Formula=$formula;NumberFormat=[string]$cell.NumberFormat;
                Font=[string]$font.Name;Size=$font.Size;Bold=$font.Bold;Italic=$font.Italic;Color=$font.Color;Fill=$fill.Color;Wrap=$cell.WrapText;
                Horizontal=$cell.HorizontalAlignment;Vertical=$cell.VerticalAlignment;Borders=$border;
                Merge=@(([int]$merge.Row-$top),([int]$merge.Column-$left),[int]$merge.Rows.Count,[int]$merge.Columns.Count)}
        } finally {Release-Com $merge;Release-Com $fill;Release-Com $font;Release-Com $cell}
    }}
    return ,$result
}
function Compare-Cells($expected,$actual,[bool]$formulas,[bool]$styles=$true) {
    $differences=@()
    for($i=0;$i -lt $expected.Count;$i++){
        $e=$expected[$i];$a=$actual[$i]
        $properties=@('Kind','Value');if($styles){$properties+='Merge'}
        # Covered positions have no independently portable style in SpreadsheetML.
        $covered=($e.Merge[2] -gt 1 -or $e.Merge[3] -gt 1) -and ($e.Row -ne $e.Merge[0] -or $e.Column -ne $e.Merge[1])
        if($styles -and -not $covered){$properties+=@('Text','Font','Size','Bold','Italic','Color','Fill','Wrap','Horizontal','Vertical','Borders')}
        foreach($key in $properties){
            $ev=$e[$key]|ConvertTo-Json -Depth 6 -Compress;$av=$a[$key]|ConvertTo-Json -Depth 6 -Compress
            if($ev -cne $av){$differences+=@{Cell=$i;Property=$key;Expected=$e[$key];Actual=$a[$key]}}
        }
        if($formulas){if(($e.Formula -replace '[()]','') -cne ($a.Formula -replace '[()]','')){$differences+=@{Cell=$i;Property='Formula';Expected=$e.Formula;Actual=$a.Formula}}}
        elseif($null -ne $a.Formula){$differences+=@{Cell=$i;Property='UnexpectedFormula';Actual=$a.Formula}}
    }
    return ,$differences
}
function Fixture($book,[bool]$relative) {
    $sheet=$book.Worksheets.Item(1);$range=$sheet.Range('C4:F9');$range.Font.Name='Arial';$range.Font.Size=12;$range.VerticalAlignment=-4108;$range.ColumnWidth=24
    $sheet.Range('C4').NumberFormat='@';$sheet.Range('C4').Value2='00123'
    $sheet.Range('D4').Value2=[double]0.12;$sheet.Range('D4').NumberFormat='0.00%'
    $sheet.Range('E4').Value2=$true;$sheet.Range('F4').Formula='=1/0'
    $sheet.Range('C5').Value2="中文`n行`t列";$sheet.Range('C5').WrapText=$true
    $sheet.Range('D5').Value2=[double]45000.5;$sheet.Range('D5').NumberFormat='yyyy-mm-dd hh:mm:ss'
    $sheet.Range('E5').Formula=$(if($relative){'=D4*4'}else{'=D4+$D$4+$D4+D$4'})
    $sheet.Range('F5').Value2=[double]45000.5;$sheet.Range('F5').NumberFormat='mmmm'
    $sheet.Range('C6:D7').Merge();$sheet.Range('C6').Value2='merged'
    $sheet.Range('C4').Font.Bold=$true;$sheet.Range('C4').Font.Color=255;$sheet.Range('C4').Interior.Color=65535
    $sheet.Range('D4').Borders.LineStyle=1
    $sheet.Range('E6').Value2=[double]0.5;$sheet.Range('E6').NumberFormat='hh:mm:ss'
    $sheet.Range('F6').Value2=[double]12.123456789;$sheet.Range('F6').NumberFormat='[h]:mm:ss'
    $sheet.Range('E7').Value2=[double]1234.56;$sheet.Range('E7').NumberFormat='"$"#,##0.00';$sheet.Range('E7').Font.Italic=$true
    $sheet.Range('F7').Value2=[double]0.00012345678;$sheet.Range('F7').NumberFormat='0.000E+00'
    $sheet.Range('C8:F8').NumberFormat='@';$sheet.Range('C8').Value2='123';$sheet.Range('D8').Value2='123456789012345678901234';$sheet.Range('E8').Value2='=1+2';$sheet.Range('F8').Value2="引号`"`r`n制表`t"
    Release-Com $range;Release-Com $sheet
}
$excel=$null;$source=$null;$dest=$null;$cases=@()
Push-Location $root
try {
    if(-not $SkipBuild){dotnet build QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release --artifacts-path artifacts/build -v minimal;if($LASTEXITCODE -ne 0){throw 'Build failed'}}
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
    foreach($case in @(
        @{Name='excel1900';Date1904=$false;TargetDate1904=$false;Mode='all'},
        @{Name='excel1904';Date1904=$true;TargetDate1904=$true;Mode='all'},
        @{Name='excel1900-to-1904';Date1904=$false;TargetDate1904=$true;Mode='all'},
        @{Name='excel1904-to-1900';Date1904=$true;TargetDate1904=$false;Mode='all'},
        @{Name='relative-formulas';Date1904=$false;TargetDate1904=$false;Mode='formulas'})) {
        $source=$excel.Workbooks.Add();$source.Date1904=$case.Date1904;Fixture $source ($case.Mode -eq 'formulas');$excel.CalculateFullRebuild()
        $source.SaveAs((Join-Path $run ($case.Name+'-source.xlsx')),51);$sheet=$source.Worksheets.Item(1);$expected=Capture $sheet 4 3
        if($case.Date1904 -ne $case.TargetDate1904) {
            # Only D5 and F5 are calendar dates. Time/duration serials must remain unchanged.
            $offset=if($case.Date1904){1462}else{-1462}
            foreach($index in 5,7){$expected[$index].Value=[double]$expected[$index].Value+$offset}
        }
        $dest=$excel.Workbooks.Add();$dest.Date1904=$case.TargetDate1904;$target=$dest.Worksheets.Item(1);$target.Range('G10:J15').ColumnWidth=24
        $range=$sheet.Range('C4:F9');$range.Copy();Release-Com $range
        $dir=Run-QuickExcel $case.Name 'import' $case.Mode
        # The QuickExcel process has exited: FlushAsync must leave formats usable by Excel.
        $target.Paste($target.Range('G10'));$excel.CalculateFullRebuild();$actual=Capture $target 10 7
        $diff=Compare-Cells $expected $actual ($case.Mode -eq 'formulas') ($case.Mode -ne 'formulas')
        if($dest.Date1904 -ne $case.TargetDate1904){$diff+=@{Property='DateSystemChanged';Expected=$case.TargetDate1904;Actual=$dest.Date1904}}
        $output=Join-Path $run ($case.Name+'-output.xlsx');$dest.SaveAs($output,51);$dest.Close($false);Release-Com $target;Release-Com $dest;$dest=$null
        $dest=$excel.Workbooks.Open($output,0,$true);$target=$dest.Worksheets.Item(1);$reopened=Capture $target 10 7
        $stable=($actual|ConvertTo-Json -Depth 10 -Compress) -ceq ($reopened|ConvertTo-Json -Depth 10 -Compress)
        $cases+=@{Name=$case.Name;Passed=($diff.Count -eq 0 -and $stable);Differences=$diff;Expected=$expected;Actual=$actual;ReopenedStable=$stable;Path=$dir}
        $dest.Close($false);Release-Com $target;Release-Com $dest;$dest=$null;$source.Close($false);Release-Com $sheet;Release-Com $source;$source=$null
        Write-Output "$($case.Name): $($diff.Count) differences; reopen=$stable"
    }
    foreach($mode in 'all','formulas') {
        $dir=Run-QuickExcel ('quickexcel-'+$mode) 'export' $mode
        $relay=Run-QuickExcel ('quickexcel-'+$mode+'-second-instance') 'import' $mode
        $beforeModel=Get-Content (Join-Path $dir 'model.json') -Raw|ConvertFrom-Json
        $relayModel=Get-Content (Join-Path $relay 'model.json') -Raw|ConvertFrom-Json
        $relayDiff=@()
        foreach($cell in $beforeModel){$found=@($relayModel|Where-Object {$_.row -eq $cell.row -and $_.column -eq $cell.column});if($found.Count -eq 0 -and $cell.kind -eq 'Blank'){continue};if($found.Count -ne 1 -or $found[0].kind -cne $cell.kind -or $found[0].value -cne $cell.value){$relayDiff+=@{Row=$cell.row;Column=$cell.column;Expected=$cell;Actual=$found}}}
        $cases+=@{Name='quickexcel-'+$mode+'-second-instance';Passed=($relayDiff.Count -eq 0);Differences=$relayDiff;Path=$dir;ReturnPath=$relay}
        Write-Output "quickexcel-$mode-second-instance: $($relayDiff.Count) differences"
        $dest=$excel.Workbooks.Add();$target=$dest.Worksheets.Item(1);$target.Range('C4:F9').ColumnWidth=24;$target.Paste($target.Range('C4'));$excel.CalculateFullRebuild()
        $actual=Capture $target 4 3;$dest.SaveAs((Join-Path $run ('quickexcel-'+$mode+'-excel.xlsx')),51)
        $range=$target.Range('C4:F9');$range.Copy();Release-Com $range
        $back=Run-QuickExcel ('quickexcel-'+$mode+'-return') 'import' $mode
        $beforeModel=Get-Content (Join-Path $dir 'model.json') -Raw|ConvertFrom-Json
        $afterModel=Get-Content (Join-Path $back 'model.json') -Raw|ConvertFrom-Json
        $diff=@()
        foreach($cell in $beforeModel){$found=@($afterModel|Where-Object {$_.row -eq $cell.row -and $_.column -eq $cell.column});if($found.Count -eq 0 -and $cell.kind -eq 'Blank'){continue};if($found.Count -ne 1 -or $found[0].kind -cne $cell.kind -or $found[0].value -cne $cell.value){$diff+=@{Row=$cell.row;Column=$cell.column;Expected=$cell;Actual=$found}}}
        $cases+=@{Name='quickexcel-'+$mode+'-roundtrip';Passed=($diff.Count -eq 0);Differences=$diff;Excel=$actual;Path=$dir;ReturnPath=$back}
        $dest.Close($false);Release-Com $target;Release-Com $dest;$dest=$null
        Write-Output "quickexcel-$mode-roundtrip: $($diff.Count) differences"
    }
    @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Passed=(@($cases|Where-Object {-not $_.Passed}).Count -eq 0);Cases=$cases;Scope='Real OS clipboard, Avalonia handlers, independent Excel COM inspection; no physical shortcut/IME claim'}|ConvertTo-Json -Depth 16|Set-Content (Join-Path $run 'report.json') -Encoding utf8
    Write-Output "Evidence: $run"
    if(@($cases|Where-Object {-not $_.Passed}).Count){throw 'Clipboard interoperability differences; inspect report.json'}
} finally {
    if($source){$source.Close($false);Release-Com $source};if($dest){$dest.Close($false);Release-Com $dest}
    if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();Pop-Location
}
