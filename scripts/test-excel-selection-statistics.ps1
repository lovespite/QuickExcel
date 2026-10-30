param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$excel=$null;$book=$null;$sheet=$null;$functions=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if(-not ('StatsExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class StatsExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
    [uint32]$ownedExcelProcess=0;[void][StatsExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
    [ordered]@{ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime;Owner='test-excel-selection-statistics.ps1'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $target 'excel-owner.json') -Encoding utf8
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Name='Statistics';$functions=$excel.WorksheetFunction
    foreach($item in @(@('A1','Amount'),@('B1','Keep'),@('A6','00123'))){$cell=$sheet.Range($item[0]);try{$cell.NumberFormat='@';$cell.Value2=[string]$item[1]}finally{Release-Com $cell}}
    foreach($item in @(@('A2',1.5),@('A3',2.5),@('A4',100),@('A5',200))){$cell=$sheet.Range($item[0]);try{$cell.Value2=[double]$item[1]}finally{Release-Com $cell}}
    foreach($item in @(@('A7','=""'),@('A8','=TRUE'),@('A9','=#N/A'),@('A11','=3'))){$cell=$sheet.Range($item[0]);try{$cell.Formula2=[string]$item[1]}finally{Release-Com $cell}}
    for($row=2;$row -le 11;$row++){$cell=$sheet.Range("B$row");try{$cell.Value2=[string]$(if($row -eq 5){'no'}else{'yes'})}finally{Release-Com $cell}}
    $records=[Collections.Generic.List[object]]::new()
    foreach($state in @('all','manual-hidden','filtered-and-manual')){
        if($state -eq 'filtered-and-manual'){$range=$sheet.Range('A1:B11');try{[void]$range.AutoFilter(2,'yes')}finally{Release-Com $range}}
        if($state -ne 'all'){$row=$sheet.Range('A4').EntireRow;try{$row.Hidden=$true}finally{Release-Com $row}}
        $excel.CalculateFullRebuild();$range=$sheet.Range('A2:A11')
        try{$records.Add([ordered]@{State=$state;Sum=$functions.Aggregate(9,7,$range);Average=$functions.Aggregate(1,7,$range);Numbers=$functions.Aggregate(2,7,$range);Count=$functions.Subtotal(103,$range)})}finally{Release-Com $range}
    }
    Release-Com $sheet;$sheet=$null;Release-Com $functions;$functions=$null
    $path=Join-Path $target 'statistics.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$excel.CalculateFullRebuild();$book.Close($false);Release-Com $book;$book=$null
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);WorkbookSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;NormalReopen=$true;Cases=$records.ToArray()}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $target 'statistics.json') -Encoding utf8
    Write-Output 'PASS: visible-row statistics compared with Excel AGGREGATE and SUBTOTAL.'
}finally{Release-Com $functions;Release-Com $sheet;if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
