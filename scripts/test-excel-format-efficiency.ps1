param([Parameter(Mandatory)][string]$SpecPath,[Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent; $target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
$specFile=[IO.Path]::GetFullPath($SpecPath,$repo);$spec=Get-Content -LiteralPath $specFile -Raw|ConvertFrom-Json
Copy-Item -LiteralPath $specFile -Destination (Join-Path $target 'spec.json')
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Capture($book){
    $records=[Collections.Generic.List[object]]::new();$sheet=$book.Worksheets.Item($(if($spec.Sheet){$spec.Sheet}else{'Precision'}))
    try{foreach($case in $spec.Cases){$cell=$sheet.Range($case.Address);try{$local=[string]$cell.NumberFormat;$records.Add([ordered]@{Id=$case.Id;Address=$case.Address;Value=$cell.Value2;Text=[string]$cell.Text;Formula=[string]$cell.Formula2;Format=$local.Replace('[红色]','[Red]').Replace('G/通用格式','General');LocalFormat=$local})}finally{Release-Com $cell}}}finally{Release-Com $sheet}
    return ,$records.ToArray()
}
$excel=$null;$book=$null;$bars=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if(-not ('FormatExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class FormatExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
    [uint32]$ownedExcelProcess=0;[void][FormatExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
    [ordered]@{ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime;Owner='test-excel-format-efficiency.ps1'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $target 'excel-owner.json') -Encoding utf8
    $sourceHash=$null
    if($WorkbookPath){$source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true)}
    else{
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$bars=$excel.CommandBars
        try{
            $sheet.Name=$(if($spec.Sheet){$spec.Sheet}else{'Precision'});[void]$sheet.Activate();$column=$sheet.Range('A:C');try{$column.ColumnWidth=60}finally{Release-Com $column}
            foreach($case in $spec.Cases){
                $cell=$sheet.Range($case.Address)
                try{
                    if($case.Formula){$cell.Formula2=[string]$case.Formula}elseif($case.Kind -eq 'Text'){$cell.NumberFormat='@';$cell.Value2=[string]$case.Value}else{$cell.Value2=[double]$case.Value}
                    $format=[string]$case.Format;if($excel.LanguageSettings.LanguageID(2) -eq 2052){$format=$format.Replace('[Red]','[红色]')}
                    if($case.Format -eq 'General'){[void]$cell.ClearFormats()}else{$cell.NumberFormat=$format};[void]$cell.Select()
                    foreach($delta in $case.Steps){$id=if($delta -gt 0){'DecimalsIncrease'}else{'DecimalsDecrease'};if(-not $bars.GetEnabledMso($id)){throw "Excel command disabled: $id"};[void]$bars.ExecuteMso($id)}
                }catch{throw "Precision case $($case.Id): $_"}finally{Release-Com $cell}
            }
        }finally{Release-Com $sheet}
    }
    $excel.CalculateFullRebuild();$initial=Capture $book;$path=Join-Path $target 'precision.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$excel.CalculateFullRebuild();$reopened=Capture $book;$book.Close($false);Release-Com $book;$book=$null
    $stable=($initial|ConvertTo-Json -Depth 5 -Compress) -ceq ($reopened|ConvertTo-Json -Depth 5 -Compress)
    $sourceUnchanged=-not $WorkbookPath -or $sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);SpecSha256=(Get-FileHash -LiteralPath $specFile -Algorithm SHA256).Hash;
        SourceSha256=$sourceHash;SourceUnchanged=$sourceUnchanged;WorkbookSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;Stable=$stable;Cases=$initial}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $target 'precision.json') -Encoding utf8
    if(-not $stable -or -not $sourceUnchanged){throw 'Precision save/reopen differs; see precision.json.'}
    Write-Output "PASS: $($initial.Count) Excel precision cases, ordinary save/reopen stable."
}finally{Release-Com $bars;if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
