param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath,[switch]$AllStyles,[switch]$EditBasics,[switch]$InspectBorders)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Color-Hex($value){$n=[int64]$value;return ('FF{0:X2}{1:X2}{2:X2}' -f ($n -band 255),(($n -shr 8)-band 255),(($n -shr 16)-band 255))}
function Capture($workbook){
 $result=[Collections.Generic.List[object]]::new()
 for($si=1;$si -le $workbook.Worksheets.Count;$si++){
  $sheet=$workbook.Worksheets.Item($si);$tables=$sheet.ListObjects
  try{for($ti=1;$ti -le $tables.Count;$ti++){
   $table=$tables.Item($ti);$range=$table.Range;$style=$table.TableStyle;$columns=$table.ListColumns
   try{
    $cols=@();for($ci=1;$ci -le $columns.Count;$ci++){$column=$columns.Item($ci);try{$cols+=@{Name=$column.Name;TotalsFunction=[int]$column.TotalsCalculation}}finally{Release-Com $column}}
    $cells=@();for($r=1;$r -le $range.Rows.Count;$r++){for($c=1;$c -le $range.Columns.Count;$c++){
     $cell=$range.Cells.Item($r,$c);$display=$cell.DisplayFormat;$font=$display.Font;$fill=$display.Interior;$row=$cell.EntireRow
     try{
      $item=@{Address=$cell.Address($false,$false);Value=$cell.Value2;Formula=[string]$cell.Formula;Text=[string]$cell.Text;Hidden=[bool]$row.Hidden;Bold=[bool]$font.Bold;Font=Color-Hex $font.Color;Fill=Color-Hex $fill.Color}
      if($InspectBorders){
       $borders=$display.Borders;$edges=@{}
       try{foreach($pair in @(@('Left',7),@('Top',8),@('Bottom',9),@('Right',10))){
        $edge=$borders.Item($pair[1]);try{
         $line=[int]$edge.LineStyle;$weight=[int]$edge.Weight
         $kind=if($line -eq -4142){0}elseif($line -eq -4119){6}elseif($line -eq -4118){4}elseif($line -eq -4115){3}elseif($line -eq 1){switch($weight){1{7};2{1};-4138{2};4{5};default{throw 'Unknown border weight'}}}else{throw 'Unknown border line'}
         $edges[$pair[0]]=@{Style=$kind;Color=if($kind){Color-Hex $edge.Color}else{$null}}
        }finally{Release-Com $edge}
       }}finally{Release-Com $borders};$item['Borders']=$edges
      }
      $cells+=$item
     }
     finally{Release-Com $row;Release-Com $fill;Release-Com $font;Release-Com $display;Release-Com $cell}
    }}
    $result.Add([ordered]@{Sheet=$sheet.Name;Name=$table.Name;Range=$range.Address($false,$false);Header=[bool]$table.ShowHeaders;Totals=[bool]$table.ShowTotals;Filter=[bool]$table.ShowAutoFilter;Style=$style.Name;First=[bool]$table.ShowTableStyleFirstColumn;Last=[bool]$table.ShowTableStyleLastColumn;Rows=[bool]$table.ShowTableStyleRowStripes;Columns=[bool]$table.ShowTableStyleColumnStripes;Headers=$cols;Cells=$cells})
   }finally{Release-Com $columns;Release-Com $style;Release-Com $range;Release-Com $table}
  }}finally{Release-Com $tables;Release-Com $sheet}
 };return ,$result.ToArray()
}
$excel=$null;$book=$null;$sourceHash=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('TableExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class TableExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
 [uint32]$ownedExcelProcess=0;[void][TableExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 $version=[ordered]@{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}
 $version|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
 if($WorkbookPath){$source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true)}
 else{
  $book=$excel.Workbooks.Add();$names=@('light','medium','dark','bands','noheader','nofilter','filtered','totals');$styles=@('TableStyleLight1','TableStyleMedium2','TableStyleDark1','TableStyleMedium9')
  if($AllStyles){$styles=@();foreach($family in @(@('Light',21),@('Medium',28),@('Dark',11))){for($s=1;$s -le $family[1];$s++){$styles+='TableStyle'+$family[0]+$s}};$names=@(1..60|ForEach-Object{'style'+$_})}
  for($i=0;$i -lt $names.Count;$i++){
   if($i -eq 0){$sheet=$book.Worksheets.Item(1)}else{$last=$book.Worksheets.Item($book.Worksheets.Count);try{$sheet=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last}}
   try{
    $sheet.Name=$names[$i];$range=$sheet.Range('B2:E7');$sheetColumns=$sheet.Range('B:E');$sheetColumns.ColumnWidth=20;Release-Com $sheetColumns
    for($c=0;$c -lt 4;$c++){$cell=$sheet.Cells.Item(2,$c+2);$cell.Value2=[string]@('Group','Amount','Code','Label')[$c];Release-Com $cell}
    for($r=0;$r -lt 5;$r++){
     $cell=$sheet.Cells.Item($r+3,2);$cell.Value2=[string]@('A','B','A','B','A')[$r];Release-Com $cell
     $cell=$sheet.Cells.Item($r+3,3);$cell.Value2=[double](($r+1)*10);Release-Com $cell
     $cell=$sheet.Cells.Item($r+3,4);$cell.NumberFormat='@';$cell.Value2='00'+($r+1);Release-Com $cell
     $cell=$sheet.Cells.Item($r+3,5);$cell.Value2='row'+($r+1);Release-Com $cell
    }
    $tables=$sheet.ListObjects;$table=$tables.Add(1,$range,[Type]::Missing,1);$table.Name='QE_'+$names[$i];$table.TableStyle=$styles[$i%$styles.Count]
    $table.ShowTableStyleFirstColumn=($AllStyles -or $i -eq 3);$table.ShowTableStyleLastColumn=($AllStyles -or $i -eq 3);$table.ShowTableStyleColumnStripes=($AllStyles -or $i -eq 3)
    if(-not $AllStyles -and $i -eq 4){$table.ShowHeaders=$false}
    if(-not $AllStyles -and $i -eq 5){$table.ShowAutoFilter=$false}
    if(-not $AllStyles -and $i -eq 6){[void]$range.AutoFilter(1,'A')}
    if($AllStyles -or $i -eq 7){$table.ShowTotals=$true;$cols=$table.ListColumns;for($c=1;$c -le 4;$c++){$col=$cols.Item($c);$col.TotalsCalculation=0;Release-Com $col};Release-Com $cols;$cell=$sheet.Range('B8');$cell.Value2='Total';Release-Com $cell;$cell=$sheet.Range('C8');$cell.Value2=[double]150;Release-Com $cell}
    Release-Com $table;Release-Com $tables;Release-Com $range
   }finally{Release-Com $sheet}
  }
 }
 if($EditBasics){
  $sheet=$book.Worksheets.Item(1);$tables=$sheet.ListObjects;$table=$tables.Item(1);$cell=$sheet.Range('C3');$range=$sheet.Range('B2:E9')
  try{$cell.Value2=[double]25;$table.Resize($range)}finally{Release-Com $range;Release-Com $cell;Release-Com $table;Release-Com $tables;Release-Com $sheet}
 }
 $excel.CalculateFullRebuild();$before=Capture $book;$path=Join-Path $target 'tables.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
 $book=$excel.Workbooks.Open($path,0,$true);$after=Capture $book;$book.Close($false);Release-Com $book;$book=$null
 $version['Before']=$before;$version['After']=$after;$version['Stable']=(($before|ConvertTo-Json -Depth 15 -Compress) -eq ($after|ConvertTo-Json -Depth 15 -Compress));$version['Sha256']=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
 if($sourceHash){$version['SourceSha256']=$sourceHash;$version['SourceUnchanged']=$sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}
 $version|ConvertTo-Json -Depth 15|Set-Content (Join-Path $target 'excel-tables.json') -Encoding utf8
}finally{if($book){$book.Close($false);Release-Com $book};if($excel){[GC]::Collect();[GC]::WaitForPendingFinalizers();$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
