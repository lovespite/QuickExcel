param([Parameter(Mandatory)][string]$SpecPath,[Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
$specFile=[IO.Path]::GetFullPath($SpecPath,$repo);$spec=Get-Content -LiteralPath $specFile -Raw|ConvertFrom-Json
Copy-Item -LiteralPath $specFile -Destination (Join-Path $target 'spec.json')
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Capture($workbook){
    $records=[Collections.Generic.List[object]]::new()
    foreach($case in $spec.Cases){
        $sheet=$workbook.Worksheets.Item([string]$case.Id)
        try{
            $hidden=@();$order=@();$texts=@()
            for($i=2;$i -le $spec.Rows.Count+1;$i++){
                $row=$sheet.Rows.Item($i);try{if($row.Hidden){$hidden+=$i}}finally{Release-Com $row}
                $cell=$sheet.Range("C$i");try{$order+=$cell.Value2}finally{Release-Com $cell}
                $cell=$sheet.Range("A$i");try{$texts+=[string]$cell.Text}finally{Release-Com $cell}
            }
            $records.Add([ordered]@{Id=$case.Id;HiddenRows=$hidden;Order=$order;Text=$texts;AutoFilterMode=[bool]$sheet.AutoFilterMode;FilterMode=[bool]$sheet.FilterMode})
        }finally{Release-Com $sheet}
    }
    return ,$records.ToArray()
}
$excel=$null;$book=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if(-not ('FilterExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class FilterExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
    [uint32]$ownedExcelProcess=0;[void][FilterExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
    [ordered]@{ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime;Owner='test-excel-filter-persistence.ps1'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $target 'excel-owner.json') -Encoding utf8
    $sourceHash=$null
    if($WorkbookPath){$source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true)}
    else{
        $book=$excel.Workbooks.Add();$index=0
        foreach($case in $spec.Cases){
            $sheet=if($index++ -eq 0){$book.Worksheets.Item(1)}else{$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))}
            try{
                $sheet.Name=[string]$case.Id;$columns=$sheet.Range('A:C');try{$columns.ColumnWidth=24}finally{Release-Com $columns}
                foreach($pair in @(@('A1','Label'),@('B1','Group'),@('C1','Identity'))){$cell=$sheet.Range($pair[0]);try{$cell.Value2=[string]$pair[1]}finally{Release-Com $cell}}
                for($r=0;$r -lt $spec.Rows.Count;$r++){
                    $item=$spec.Rows[$r];$row=$r+2;$cell=$sheet.Range("A$row")
                    try{
                        switch($item.Kind){'Text'{$cell.NumberFormat='@';$cell.Value2=[string]$item.Value} 'Number'{$cell.Value2=[double]$item.Value} 'Formula'{$cell.Formula2=[string]$item.Value}}
                        if($item.Format){$cell.NumberFormat=[string]$item.Format}
                        if($r % 3 -eq 0){$color=$cell.Interior;try{$color.Color=255}finally{Release-Com $color}}
                        if($r % 4 -eq 0){$font=$cell.Font;try{$font.Color=16711680}finally{Release-Com $font}}
                    }finally{Release-Com $cell}
                    $cell=$sheet.Range("B$row");try{$cell.Value2=[double](($r % 6)+1)}finally{Release-Com $cell}
                    $cell=$sheet.Range("C$row");try{$cell.Value2=[double]($r+1)}finally{Release-Com $cell}
                }
                $range=$sheet.Range('A1:C'+($spec.Rows.Count+1));[void]$range.AutoFilter();$excel.CalculateFullRebuild()
                try{
                    foreach($filter in $case.Filters){
                        $criteria=if($filter.Criteria1 -is [array]){[object[]]$filter.Criteria1}else{[string]$filter.Criteria1}
                        $operator=if($filter.Operator){[int]$filter.Operator}else{0}
                        if($filter.Criteria2){[void]$range.AutoFilter([int]$filter.Field,$criteria,$operator,[string]$filter.Criteria2)}elseif($operator){[void]$range.AutoFilter([int]$filter.Field,$criteria,$operator)}else{[void]$range.AutoFilter([int]$filter.Field,$criteria)}
                    }
                    if($case.Sort){
                        $sort=$sheet.Sort;$fields=$sort.SortFields
                        try{
                            $fields.Clear()
                            foreach($key in $case.Sort){
                                $letter=[char](64+[int]$key.Column);$keyRange=$sheet.Range("$($letter)2:$($letter)$($spec.Rows.Count+1)")
                                try{
                                    $basis=if($key.Basis){[int]$key.Basis}else{0};$custom=if($key.CustomOrder){[string]$key.CustomOrder}else{[Type]::Missing}
                                    $field=$fields.Add($keyRange,$basis,[int]$key.Order,$custom)
                                    try{if($basis -ne 0){$on=$field.SortOnValue;try{$on.Color=[int]$key.Color}finally{Release-Com $on}}}finally{Release-Com $field}
                                }finally{Release-Com $keyRange}
                            }
                            $sort.SetRange($range);$sort.Header=1;$sort.MatchCase=$false;$sort.Orientation=1;$sort.Apply()
                        }finally{Release-Com $fields;Release-Com $sort}
                    }
                }finally{Release-Com $range}
                foreach($rowIndex in $case.ManualRows){$row=$sheet.Rows.Item([int]$rowIndex);try{$row.Hidden=$true}finally{Release-Com $row}}
            }catch{throw "Excel filter case $($case.Id): $_"}finally{Release-Com $sheet}
        }
    }
    $excel.CalculateFullRebuild();$initial=Capture $book;$path=Join-Path $target 'filters.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$excel.CalculateFullRebuild();$reopened=Capture $book
    $stable=($initial|ConvertTo-Json -Depth 6 -Compress) -ceq ($reopened|ConvertTo-Json -Depth 6 -Compress)
    foreach($case in $spec.Cases){$sheet=$book.Worksheets.Item([string]$case.Id);try{if($sheet.FilterMode){$filter=$sheet.AutoFilter;try{$filter.ApplyFilter()}finally{Release-Com $filter}}}finally{Release-Com $sheet}}
    $reapplied=Capture $book;$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($path,0,$true)
    $sheet=$book.Worksheets.Item('multi');try{
        $cell=$sheet.Range('B2');try{$cell.Value2=[double]6}finally{Release-Com $cell}
        $filter=$sheet.AutoFilter;try{$filter.ApplyFilter()}finally{Release-Com $filter}
    }finally{Release-Com $sheet}
    $edited=Capture $book;$editedPath=Join-Path $target 'filters-edited.xlsx';$book.SaveAs($editedPath,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($editedPath,0,$true);$editedReopened=Capture $book;$book.Close($false);Release-Com $book;$book=$null
    $editedStable=($edited|ConvertTo-Json -Depth 6 -Compress) -ceq ($editedReopened|ConvertTo-Json -Depth 6 -Compress)
    $sourceUnchanged=-not $WorkbookPath -or $sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);SpecSha256=(Get-FileHash -LiteralPath $specFile -Algorithm SHA256).Hash;
        SourceSha256=$sourceHash;SourceUnchanged=$sourceUnchanged;WorkbookSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;Stable=$stable;EditedStable=$editedStable;Cases=$initial;Reapplied=$reapplied;Edited=$edited}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $target 'filters.json') -Encoding utf8
    if(-not $stable -or -not $editedStable -or -not $sourceUnchanged){throw 'Filter save/reopen differs.'}
    Write-Output "PASS: $($initial.Count) Excel filter/sort cases; normal save/reopen stable."
}finally{if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
