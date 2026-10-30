param([Parameter(Mandatory)][string]$OutputDirectory, [string]$WorkbookPath, [string]$ReferencePath,
    [ValidateSet('AfterReopen','AfterRowInsert')][string]$ReferenceState='AfterReopen')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($object){if($null -ne $object){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}}
function Capture-Views($book,$window){
    $active=$book.ActiveSheet; try {$activeName=[string]$active.Name} finally {Release-Com $active}
    $rows=@()
    for($i=1;$i -le $book.Worksheets.Count;$i++){
        $sheet=$book.Worksheets.Item($i)
        try{
            $record=[ordered]@{Name=[string]$sheet.Name;Visible=[int]$sheet.Visible}
            if($sheet.Visible -eq -1){
                $sheet.Activate();$cell=$window.ActiveCell
                try{
                    $record.Freeze=[bool]$window.FreezePanes;$record.Rows=[int]$window.SplitRow;$record.Columns=[int]$window.SplitColumn
                    $record.Zoom=[int]$window.Zoom;$record.ActiveCell=[string]$cell.Address($false,$false)
                    $record.ScrollRow=[int]$window.ScrollRow;$record.ScrollColumn=[int]$window.ScrollColumn
                }finally{Release-Com $cell}
            }
            $rows += $record
        }finally{Release-Com $sheet}
    }
    $active=$book.Worksheets.Item($activeName);try{$active.Activate()}finally{Release-Com $active}
    return [ordered]@{ActiveSheet=$activeName;Sheets=$rows}
}
$excel=$null;$book=$null;$window=$null
try{
    $sourceHash=if($WorkbookPath){(Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$null}
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if($WorkbookPath){$book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath,$repo),0,$true)}
    else{
        $book=$excel.Workbooks.Add()
        while($book.Worksheets.Count -gt 1){$extra=$book.Worksheets.Item($book.Worksheets.Count);try{$extra.Delete()}finally{Release-Com $extra}}
        $names=@('普通','冻结双轴','冻结行','冻结列','隐藏','非常隐藏')
        $window=$book.Windows.Item(1)
        for($i=0;$i -lt $names.Count;$i++){
            if($i -eq 0){$sheet=$book.Worksheets.Item(1)}else{
                $last=$book.Worksheets.Item($book.Worksheets.Count)
                try{$sheet=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last}
            }
            try{
                $sheet.Name=$names[$i];$sheet.Activate()
                $cell=$sheet.Cells.Item(1,1);try{$cell.Value2='view fixture'}finally{Release-Com $cell}
                if($i -lt 4){
                    $row=@(0,2,1,0)[$i];$column=@(0,1,0,2)[$i]
                    $window.FreezePanes=$false;$window.SplitRow=$row;$window.SplitColumn=$column
                    if($row+$column -gt 0){$window.FreezePanes=$true}
                    $window.Zoom=@(90,150,125,80)[$i]
                    $cell=$sheet.Range(@('D4','D6','C9','F7')[$i]);try{$cell.Select()}finally{Release-Com $cell}
                    $window.ScrollRow=@(1,6,4,2)[$i];$window.ScrollColumn=@(1,3,1,4)[$i]
                }else{$sheet.Visible=if($i -eq 4){0}else{2}}
            }finally{Release-Com $sheet}
        }
        $active=$book.Worksheets.Item('冻结双轴');try{$active.Activate()}finally{Release-Com $active}
    }
    if(-not $window){$window=$book.Windows.Item(1)}
    $before=Capture-Views $book $window
    $referenceMatches=$null
    if($ReferencePath){
        $reference=(Get-Content -LiteralPath ([IO.Path]::GetFullPath($ReferencePath,$repo)) -Raw|ConvertFrom-Json).$ReferenceState
        $referenceMatches=($reference|ConvertTo-Json -Depth 8 -Compress) -eq ($before|ConvertTo-Json -Depth 8 -Compress)
    }
    $savedPath=Join-Path $target $(if($WorkbookPath){'excel-resaved.xlsx'}else{'excel-source.xlsx'})
    $book.SaveAs($savedPath,51)
    $calculation=$excel.Calculation
    Release-Com $window;$window=$null;$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($savedPath,0,$true);$window=$book.Windows.Item(1)
    $after=Capture-Views $book $window
    $same=($before|ConvertTo-Json -Depth 8 -Compress) -eq ($after|ConvertTo-Json -Depth 8 -Compress)
    $afterInsert=$null
    if(-not $WorkbookPath){
        $sheet=$book.Worksheets.Item('冻结双轴')
        try{$sheet.Activate();$row=$sheet.Rows.Item(1);try{$row.Insert()}finally{Release-Com $row}}
        finally{Release-Com $sheet}
        $afterInsert=Capture-Views $book $window
        $book.SaveAs((Join-Path $target 'excel-row-inserted.xlsx'),51)
    }
    Release-Com $window;$window=$null;$book.Close($false);Release-Com $book;$book=$null
    $unchanged= -not $WorkbookPath -or $sourceHash -eq (Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash
    [ordered]@{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=$calculation;
        CorruptLoad='xlNormalLoad';SourceUnchanged=$unchanged;ReferenceMatches=$referenceMatches;SameAfterReopen=$same;
        BeforeSave=$before;AfterReopen=$after;AfterRowInsert=$afterInsert;SavedSha256=(Get-FileHash -LiteralPath $savedPath).Hash}|
        ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $target 'excel-views.json') -Encoding utf8
    if(-not $same -or -not $unchanged -or $referenceMatches -eq $false){throw 'Excel view comparison failed; see excel-views.json.'}
    Write-Output $target
}finally{
    Release-Com $window
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
}
