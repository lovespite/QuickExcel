param([Parameter(Mandatory)][string]$SpecPath,[Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
$specFile=[IO.Path]::GetFullPath($SpecPath,$repo);$spec=Get-Content -LiteralPath $specFile -Raw|ConvertFrom-Json
Copy-Item -LiteralPath $specFile -Destination (Join-Path $target 'spec.json')
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$excel=$null;$book=$null;$sheet=$null;$cell=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if(-not ('ArityExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ArityExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
    [uint32]$ownedExcelProcess=0;[void][ArityExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
    [ordered]@{ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime;Owner='test-excel-formula-arity.ps1'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $target 'excel-owner.json') -Encoding utf8
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$records=[Collections.Generic.List[object]]::new();$row=0
    foreach($case in $spec.Cases){
        $row++;$cell=$sheet.Range("A$row");$accepted=$true;$reason=$null;$formula=$null
        try{$cell.Formula2=[string]$case.Formula;$formula=[string]$cell.Formula2}catch{$accepted=$false;$reason=$_.Exception.Message}finally{Release-Com $cell;$cell=$null}
        $expected=if($case.PSObject.Properties.Name -contains 'ExcelAccepted'){$case.ExcelAccepted}else{$case.ExpectedAccepted}
        $records.Add([ordered]@{Function=$case.Function;Count=$case.Count;Input=$case.Formula;Accepted=$accepted;Formula=$formula;Rejection=$reason;ExpectedAccepted=$expected;Matches=$accepted -eq $expected})
    }
    $excel.CalculateFullRebuild();Release-Com $sheet;$sheet=$null
    $path=Join-Path $target 'accepted.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$excel.CalculateFullRebuild();$book.Close($false);Release-Com $book;$book=$null
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);SpecSha256=(Get-FileHash -LiteralPath $specFile -Algorithm SHA256).Hash;
        WorkbookSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;NormalReopen=$true;Cases=$records.ToArray()}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $target 'arity.json') -Encoding utf8
    Write-Output "Arity cases: $($records.Count); registry differences: $(@($records|Where-Object{-not $_.Matches}).Count)"
    if(@($records|Where-Object{-not $_.Matches}).Count -ne 0){throw 'Excel arity differs from the declared oracle; see arity.json.'}
}finally{Release-Com $cell;Release-Com $sheet;if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
