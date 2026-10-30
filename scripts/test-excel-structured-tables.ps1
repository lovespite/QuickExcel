param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath,[switch]$DeleteAllData)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Text($sheet,$address,[string]$value){$cell=$sheet.Range($address);try{$cell.Value2=$value}finally{Release-Com $cell}}
function Number($sheet,$address,[double]$value){$cell=$sheet.Range($address);try{$cell.Value2=$value}finally{Release-Com $cell}}
function Formula($sheet,$address,[string]$value){$cell=$sheet.Range($address);try{$cell.Formula=$value}finally{Release-Com $cell}}
function Capture($workbook){
 $records=@()
 for($si=1;$si -le $workbook.Worksheets.Count;$si++){
  $sheet=$workbook.Worksheets.Item($si);$used=$sheet.UsedRange;$objects=$sheet.ListObjects
  try{
   $cells=@();$tables=@()
   for($r=1;$r -le [Math]::Min(30,$used.Row+$used.Rows.Count);$r++){for($c=1;$c -le [Math]::Min(15,$used.Column+$used.Columns.Count);$c++){
    $cell=$sheet.Cells.Item($r,$c)
    try{
     $value=$cell.Value2;$has=[bool]$cell.HasFormula;if($null -eq $value -and -not $has){continue}
     $kind=if($value -is [string]){'Text'}elseif($value -is [bool]){'Boolean'}elseif($null -eq $value){'Blank'}else{'Number'}
     if($has -and [string]$cell.Text -in @('#REF!','#VALUE!','#DIV/0!','#N/A','#NAME?','#NUM!','#NULL!','#SPILL!')){$kind='Error';$value=[string]$cell.Text}
     $cells+=@{Address=$cell.Address($false,$false);Kind=$kind;Value=$value;Formula=if($has){[string]$cell.Formula}else{$null}}
    }finally{Release-Com $cell}
   }}
   for($ti=1;$ti -le $objects.Count;$ti++){
    $table=$objects.Item($ti);$range=$table.Range;$columns=$table.ListColumns
    try{$names=@();for($ci=1;$ci -le $columns.Count;$ci++){$col=$columns.Item($ci);try{$names+=$col.Name}finally{Release-Com $col}};$tables+=@{Name=$table.Name;Range=$range.Address($false,$false);Headers=$names;Totals=[bool]$table.ShowTotals}}
    finally{Release-Com $columns;Release-Com $range;Release-Com $table}
   }
   $records+=@{Sheet=$sheet.Name;Cells=$cells;Tables=$tables}
  }finally{Release-Com $objects;Release-Com $used;Release-Com $sheet}
 };return ,$records
}
function Snapshot([string]$name){
 $excel.CalculateFullRebuild();$before=Capture $script:book;$path=Join-Path $target ($name+'.xlsx')
 $script:book.SaveAs($path,51);$script:book.Close($false);Release-Com $script:book;$script:book=$null
 $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
 $script:book=$excel.Workbooks.Open($path,0,$false);$excel.CalculateFullRebuild();$after=Capture $script:book
 $record=@{Name=$name;Path=$path;Sha256=$hash;Before=$before;After=$after;Stable=(($before|ConvertTo-Json -Depth 15 -Compress)-eq($after|ConvertTo-Json -Depth 15 -Compress))}
 $record|ConvertTo-Json -Depth 15|Set-Content (Join-Path $target ($name+'.json')) -Encoding utf8
 if(-not $record.Stable){throw ('Unstable stage: '+$name)}
}
function Sheet([string]$name){$last=$book.Worksheets.Item($book.Worksheets.Count);try{$sheet=$book.Worksheets.Add([Type]::Missing,$last);$sheet.Name=$name;return $sheet}finally{Release-Com $last}}
function Table($sheet,[string]$range,[string]$name){$cells=$sheet.Range($range);$tables=$sheet.ListObjects;try{$table=$tables.Add(1,$cells,[Type]::Missing,1);$table.Name=$name;return $table}finally{Release-Com $tables;Release-Com $cells}}
$excel=$null;$script:book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('StructuredExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class StructuredExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
 [uint32]$ownedExcelProcess=0;[void][StructuredExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
 if($WorkbookPath){
  $source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$script:book=$excel.Workbooks.Open($source,0,$true)
  if($DeleteAllData){$sheet=$book.Worksheets.Item('A1Column');$range=$sheet.Range('A2:C3');$rows=$range.EntireRow;try{[void]$rows.Delete()}finally{Release-Com $rows;Release-Com $range;Release-Com $sheet}}
  Snapshot 'output'
  @{Sha256=$hash;Unchanged=$hash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content (Join-Path $target 'source.json');return
 }
 $script:book=$excel.Workbooks.Add(-4167);$data=$book.Worksheets.Item(1);$data.Name='Data'
 try{
  foreach($item in @(@('B2','Qty'),@('C2','Price'),@('D2','Amount'),@('E2','Tax'))){Text $data $item[0] $item[1]}
  for($r=3;$r -le 5;$r++){Number $data "B$r" ($r-1);Number $data "C$r" (10*($r-2))}
  $table=Table $data 'B2:E5' 'Sales'
  try{Formula $data 'D3:D5' '=[@Qty]*[@Price]';Formula $data 'E3:E5' '=[@Amount]*0.1';$table.ShowTotals=$true
   $columns=$table.ListColumns;try{for($c=1;$c -le 4;$c++){$column=$columns.Item($c);try{$column.TotalsCalculation=0}finally{Release-Com $column}}}finally{Release-Com $columns}
   Formula $data 'D6' '=SUM(Sales[Amount])'
  }finally{Release-Com $table}
 }finally{Release-Com $data}
 $summary=Sheet 'Summary'
 try{
  $formulas=@('=SUM(Sales[Amount])','=SUM(Sales)','=SUM(Sales[[#Headers],[#Data],[Amount]])','=SUM(Sales[[#All],[Amount]])',
   '=Sales[[#Headers],[Price]]','=Sales[[#Totals],[Amount]]','=SUM(Sales[[Qty]:[Amount]])','=SUM(Sales[[#Data],[Qty]:[Price]])',
   '=INDEX(Sales[Amount],2)','=SUM(Sales[#Headers])','=SUM(Sales[#Totals])','=SUM(SalesAmount)')
  $names=$book.Names;try{$name=$names.Add('SalesAmount','=Sales[Amount]');Release-Com $name}finally{Release-Com $names}
  for($i=0;$i -lt $formulas.Count;$i++){Formula $summary ('A'+($i+1)) $formulas[$i]}
  for($r=1;$r -le 7;$r++){Formula $summary "B$r" '=Sales[@Amount]'}
 }finally{Release-Com $summary}
 $a1=Sheet 'A1Column'
 try{Text $a1 'A1' 'Qty';Text $a1 'B1' 'Price';Text $a1 'C1' 'Amount';Number $a1 'A2' 2;Number $a1 'A3' 3;Number $a1 'B2' 4;Number $a1 'B3' 5
  $table=Table $a1 'A1:C3' 'Plain';try{Formula $a1 'C2' '=A2*B2';Formula $a1 'C3' '=A3*B3'}finally{Release-Com $table}
 }finally{Release-Com $a1}
 $special=Sheet 'Special'
 try{
  Text $special 'A1' '#Tag';Text $special 'B1' '@Value';Text $special 'C1' '[net]';Text $special 'D1' "O'Neil"
  for($r=2;$r -le 3;$r++){for($c=1;$c -le 4;$c++){$cell=$special.Cells.Item($r,$c);try{$cell.Value2=[double]($r*$c)}finally{Release-Com $cell}}}
  $table=Table $special 'A1:D3' 'Escapes';Release-Com $table
  Formula $special 'F2' "=SUM(Escapes['#Tag])";Formula $special 'F3' "=SUM(Escapes['@Value])";Formula $special 'F4' "=SUM(Escapes['[net']])";Formula $special 'F5' "=SUM(Escapes[O''Neil])"
 }finally{Release-Com $special}
 $edge=Sheet 'Boundaries'
 try{
  Text $edge 'A1' 'Qty';Text $edge 'B1' 'Amount';$table=Table $edge 'A1:B2' 'EmptyTable';$rows=$table.ListRows;$row=$rows.Item(1)
  try{$row.Delete()}finally{Release-Com $row;Release-Com $rows;Release-Com $table}
  $cases=@('=SUM(EmptyTable[Qty])','=COUNT(EmptyTable[Qty])','=COUNTA(EmptyTable[Qty])','=AVERAGE(EmptyTable[Qty])','=INDEX(EmptyTable[Qty],1)',
   '=SUM(EmptyTable[#Totals])','=SUM(EmptyTable[#Headers])','=SUM(EmptyTable[#All])')
  for($i=0;$i -lt $cases.Count;$i++){Formula $edge ('D'+($i+1)) $cases[$i]}
 }finally{Release-Com $edge}
 Snapshot 'initial'
 $data=$book.Worksheets.Item('Data');$objects=$data.ListObjects;$table=$objects.Item(1);$columns=$table.ListColumns;$column=$columns.Item(2)
 try{$column.Name='Unit Price';$table.Name='Revenue'}finally{Release-Com $column;Release-Com $columns;Release-Com $table;Release-Com $objects;Release-Com $data}
 Snapshot 'rename'
 $data=$book.Worksheets.Item('Data');$rows=$data.Rows;$row=$rows.Item(4);try{[void]$row.Insert()}finally{Release-Com $row;Release-Com $rows;Release-Com $data}
 Snapshot 'insert-row'
 $data=$book.Worksheets.Item('Data');$columns=$data.Columns;$column=$columns.Item(3);try{[void]$column.Insert()}finally{Release-Com $column;Release-Com $columns;Release-Com $data}
 Snapshot 'insert-column'
 $data=$book.Worksheets.Item('Data');$columns=$data.Columns;$column=$columns.Item(3);try{[void]$column.Delete()}finally{Release-Com $column;Release-Com $columns;Release-Com $data}
 $data=$book.Worksheets.Item('Data');$rows=$data.Rows;$row=$rows.Item(4);try{[void]$row.Delete()}finally{Release-Com $row;Release-Com $rows;Release-Com $data}
 Snapshot 'delete-inserted'
 $data=$book.Worksheets.Item('Data');$last=$book.Worksheets.Item($book.Worksheets.Count)
 try{$data.Copy([Type]::Missing,$last)}finally{Release-Com $last;Release-Com $data}
 $copy=$book.Worksheets.Item($book.Worksheets.Count);$copy.Name='Copy';$objects=$copy.ListObjects;$table=$objects.Item(1)
 try{$table.Name='Revenue_Copy';Number $copy 'B3' 7}finally{Release-Com $table;Release-Com $objects;Release-Com $copy}
 Snapshot 'copy-sheet'
 $data=$book.Worksheets.Item('Data');$summary=$book.Worksheets.Item('Summary');$source=$data.Range('D3');$destination=$summary.Range('D12');$same=$data.Range('K3')
 try{$source.Copy($destination);$source.Copy($same)}finally{Release-Com $same;Release-Com $destination;Release-Com $source;Release-Com $summary;Release-Com $data}
 Snapshot 'copy-formula'
 $a1=$book.Worksheets.Item('A1Column');$objects=$a1.ListObjects;$table=$objects.Item(1);$range=$a1.Range('A1:C5')
 try{$table.Resize($range)}finally{Release-Com $range;Release-Com $table;Release-Com $objects;Release-Com $a1}
 Snapshot 'expand'
 $a1=$book.Worksheets.Item('A1Column');$rows=$a1.Rows;$row=$rows.Item(2);try{[void]$row.Insert()}finally{Release-Com $row;Release-Com $rows;Release-Com $a1}
 Snapshot 'insert-first'
 $data=$book.Worksheets.Item('Data');$columns=$data.Columns;$column=$columns.Item(3);try{[void]$column.Delete()}finally{Release-Com $column;Release-Com $columns;Release-Com $data}
 Snapshot 'delete-column'
}finally{if($script:book){$script:book.Close($false);Release-Com $script:book};if($excel){[GC]::Collect();[GC]::WaitForPendingFinalizers();$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
