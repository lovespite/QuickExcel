param([Parameter(Mandatory)][string]$OutputDirectory, [string]$WorkbookPath, [string]$OraclePath, [ValidateSet('Initial','Edited')][string]$State='Initial', [switch]$ExtendedAggregations)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$artifactRoot = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
$target = [IO.Path]::GetFullPath($OutputDirectory,$repo)
if (-not $target.StartsWith($artifactRoot,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
if ($WorkbookPath) {
    $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath,$repo)
    if (-not $WorkbookPath.StartsWith($artifactRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Verification accepts test workbooks under artifacts only.' }
    $sourceHash=(Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($value) { if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
function Capture-Pivots($book) {
    $result = [Collections.Generic.List[object]]::new()
    foreach ($index in 1..$book.Worksheets.Count) {
        $sheet = $book.Worksheets.Item($index); $pivots = $sheet.PivotTables()
        try {
            for ($p = 1; $p -le $pivots.Count; $p++) {
                $pivot = $pivots.Item($p); $range = $pivot.TableRange2; $cache = $pivot.PivotCache()
                try {
                    $values = [Collections.Generic.List[object]]::new()
                    for ($row = 1; $row -le $range.Rows.Count; $row++) {
                        for ($col = 1; $col -le $range.Columns.Count; $col++) {
                            $cell = $range.Cells.Item($row,$col)
                            try {
                                $value = $cell.Value2
                                $kind = if ($null -eq $value) { 'Blank' } elseif ($value -is [string]) { 'Text' } elseif ($value -is [bool]) { 'Boolean' } elseif ($value -is [int] -and $value -lt -2146820000) { 'Error' } else { 'Number' }
                                if ($kind -eq 'Error') { $value=[string]$cell.Text }
                                $values.Add([ordered]@{ Row=$row-1; Column=$col-1; Kind=$kind; Value=$value })
                            } finally { Release-Com $cell }
                        }
                    }
                    $result.Add([pscustomobject][ordered]@{ Name=[string]$pivot.Name; Sheet=[string]$sheet.Name; Range=[string]$range.Address($false,$false); Source=[string]$cache.SourceData; Values=$values.ToArray() })
                } finally { Release-Com $cache; Release-Com $range; Release-Com $pivot }
            }
        } finally { Release-Com $pivots; Release-Com $sheet }
    }
    return @($result.ToArray() | Sort-Object Sheet,Name)
}
function Refresh-LocalPivots($book) {
    $caches = $book.PivotCaches()
    try {
        for ($index=1; $index -le $caches.Count; $index++) {
            $cache = $caches.Item($index)
            try {
                if ($cache.SourceType -ne 1 -or [string]$cache.SourceData -match '\[|https?:|\\\\') { throw 'Only workbook-local worksheet pivot sources may refresh in this test.' }
                $cache.Refresh()
            } finally { Release-Com $cache }
        }
    } finally { Release-Com $caches }
}
$excel=$null; $book=$null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3; $excel.AskToUpdateLinks=$false
    [ordered]@{ Hwnd=$excel.Hwnd; Version=[string]$excel.Version; Build=[string]$excel.Build } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $target 'excel-owner.json') -Encoding utf8
    if ($WorkbookPath) { $book=$excel.Workbooks.Open($WorkbookPath,0,$true) }
    else {
        $book=$excel.Workbooks.Add()
        while ($book.Worksheets.Count -gt 1) { $extra=$book.Worksheets.Item($book.Worksheets.Count); $extra.Delete(); Release-Com $extra }
        $data=$book.Worksheets.Item(1); $data.Name='Data'; $summary=$book.Worksheets.Add([Type]::Missing,$data); $summary.Name='Summary'
        try {
            $records=@(@('Region','Product','Amount'),@('North','A',10),@('South','A',20),@('North','B',5),@('South','B',0),@('North','A',7),@('South','A','text'),@('East','B',-3),@('East','A',$null))
            if ($ExtendedAggregations) { $records+=@(@('South','C','123'),@('North','C',$true),@('East','C',$null)) }
            for ($r=0; $r -lt $records.Count; $r++) { for ($c=0; $c -lt 3; $c++) {
                $cell=$data.Cells.Item($r+1,$c+1)
                try { if ($null -ne $records[$r][$c]) { if ($records[$r][$c] -is [int]) {$cell.Value2=[double]$records[$r][$c]} elseif ($records[$r][$c] -is [bool]) {$cell.Value2=$records[$r][$c]} else {if ([string]$records[$r][$c] -match '^\d+$') {$cell.NumberFormat='@'}; $cell.Value2=[string]$records[$r][$c]} } } finally { Release-Com $cell }
            } }
            $source=$data.Range('A1:C'+$records.Count); $tables=$data.ListObjects; $table=$tables.Add(1,$source,[Type]::Missing,1); $table.Name='SalesData'
            Release-Com $table; Release-Com $tables; Release-Com $source
            $caches=$book.PivotCaches(); $cache=$caches.Create(1,'SalesData',6); $pivots=$summary.PivotTables()
            try {
                $specifications=@(@('SalesSum','A3',-4157,'Total'),@('SalesCount','A12',-4112,'Count'))
                if ($ExtendedAggregations) { $specifications+=@(@('SalesAverage','A21',-4106,'Average'),@('SalesMin','A30',-4139,'Min'),@('SalesMax','A39',-4136,'Max')) }
                foreach ($spec in $specifications) {
                    $destination=$summary.Range($spec[1]); $pivot=$pivots.Add($cache,$destination,$spec[0]); Release-Com $destination
                    try {
                        $fields=$pivot.PivotFields(); $rows=$fields.Item('Region'); $columns=$fields.Item('Product'); $amount=$fields.Item('Amount')
                        try {
                            $rows.Orientation=1; $rows.Position=1; $columns.Orientation=2; $columns.Position=1
                            $valueField=$pivot.AddDataField($amount,$spec[3],$spec[2]); Release-Com $valueField
                            $rows.Subtotals=@($false)*12; $columns.Subtotals=@($false)*12
                            $pivot.RowAxisLayout(1); $pivot.RowGrand=$true; $pivot.ColumnGrand=$true
                            $pivot.DisplayFieldCaptions=$true; $pivot.TableStyle2=''; $pivot.RefreshTable() | Out-Null
                        } finally { Release-Com $amount; Release-Com $columns; Release-Com $rows; Release-Com $fields }
                    } finally { Release-Com $pivot }
                }
            } finally { Release-Com $pivots; Release-Com $cache; Release-Com $caches }
        } finally { Release-Com $summary; Release-Com $data }
    }
    $before=Capture-Pivots $book
    Refresh-LocalPivots $book
    $initial=Capture-Pivots $book
    $matches=$null
    if ($OraclePath) {
        $oracle=Get-Content -LiteralPath ([IO.Path]::GetFullPath($OraclePath,$repo)) -Raw | ConvertFrom-Json
        $matches=($oracle.$State|ConvertTo-Json -Depth 7 -Compress) -ceq ($before|ConvertTo-Json -Depth 7 -Compress) -and
            ($oracle.$State|ConvertTo-Json -Depth 7 -Compress) -ceq ($initial|ConvertTo-Json -Depth 7 -Compress)
    }
    $path=Join-Path $target 'initial.xlsx'; $book.SaveAs($path,51); $book.Close($false); Release-Com $book; $book=$null
    $book=$excel.Workbooks.Open($path,0,$true); $reopened=Capture-Pivots $book
    $stable=($initial|ConvertTo-Json -Depth 7 -Compress) -ceq ($reopened|ConvertTo-Json -Depth 7 -Compress)
    $edited=$null
    if (-not $WorkbookPath) {
        $data=$book.Worksheets.Item('Data'); $cell=$data.Range('C2'); try { $cell.Value2=12.0 } finally { Release-Com $cell; Release-Com $data }
        Refresh-LocalPivots $book; $edited=Capture-Pivots $book; $book.SaveAs((Join-Path $target 'edited.xlsx'),51)
    }
    $unchanged=if($WorkbookPath){$sourceHash -eq (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash}else{$true}
    [ordered]@{ Version=[string]$excel.Version; Build=[string]$excel.Build; Stable=$stable; ReferenceMatches=$matches; SourceUnchanged=$unchanged; BeforeRefresh=$before; Initial=$initial; Reopened=$reopened; Edited=$edited } |
        ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $target 'pivots.json') -Encoding utf8
    if (-not $stable -or $matches -eq $false -or -not $unchanged) { throw 'Pivot output changed or differs from independent oracle; inspect evidence.' }
    Write-Output "PASS: $($initial.Count) native Excel pivot tables; cached display and save/reopen stable."
} finally { if ($book) { $book.Close($false); Release-Com $book }; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); if ($excel) { $excel.Quit(); Release-Com $excel } }
