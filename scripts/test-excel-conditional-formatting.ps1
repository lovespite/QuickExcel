param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Set-Value($cell,$value){
 if($null -eq $value){[void]$cell.ClearContents()}
 elseif($value -is [string]){$cell.NumberFormat='@';$cell.Value2=[string]$value}
 elseif($value -is [bool]){$cell.Value2=[bool]$value}
 else{$cell.Value2=[double]$value}
}
function Color($value){$v=[int64]$value;return ('FF{0:X2}{1:X2}{2:X2}' -f ($v -band 255),(($v -shr 8) -band 255),(($v -shr 16) -band 255))}
function Sheet([string]$name){$last=$book.Worksheets.Item($book.Worksheets.Count);try{$s=$book.Worksheets.Add([Type]::Missing,$last);$s.Name=$name;return $s}finally{Release-Com $last}}
function Values($s,[int]$column,$values){for($i=0;$i -lt $values.Count;$i++){$c=$s.Cells.Item($i+2,$column);try{Set-Value $c $values[$i]}finally{Release-Com $c}}}
function CellRule($s,[string]$range,[int]$operator,[string]$first,[string]$second='5'){
 $cells=$s.Range($range);$rules=$cells.FormatConditions
 try{return $rules.Add(1,$operator,$first,$second)}finally{Release-Com $rules;Release-Com $cells}
}
function Expression($s,[string]$range,[string]$formula){
 $cells=$s.Range($range);$rules=$cells.FormatConditions
 try{return $rules.Add(2,[Type]::Missing,$formula)}finally{Release-Com $rules;Release-Com $cells}
}
function Paint($rule,[int]$fill,[int]$color=0,[bool]$bold=$false){
 $f=$rule.Font;$b=$rule.Interior;try{$b.Color=$fill;$f.Color=$color;$f.Bold=$bold;$rule.StopIfTrue=$false}finally{Release-Com $b;Release-Com $f}
}
function Capture($workbook){
 $result=@()
 for($si=1;$si -le $workbook.Worksheets.Count;$si++){
  $s=$workbook.Worksheets.Item($si);$used=$s.UsedRange;$rules=$used.FormatConditions
  try{
   $cells=@();$formats=@()
   for($i=1;$i -le $rules.Count;$i++){
    $rule=$rules.Item($i);$range=$rule.AppliesTo
    try{
     $f1=$null;$f2=$null;$op=$null;$text=$null
     try{$f1=$rule.Formula1}catch{};try{$f2=$rule.Formula2}catch{};try{$op=$rule.Operator}catch{};try{$text=$rule.Text}catch{}
     $formats+=@{Type=$rule.Type;Range=$range.Address($false,$false);Priority=$rule.Priority;Stop=[bool]$rule.StopIfTrue;Formula1=$f1;Formula2=$f2;Operator=$op;Text=$text}
    }finally{Release-Com $range;Release-Com $rule}
   }
   for($r=1;$r -le [Math]::Min(50,$used.Row+$used.Rows.Count);$r++){for($c=1;$c -le [Math]::Min(16,$used.Column+$used.Columns.Count);$c++){
    $cell=$s.Cells.Item($r,$c);$display=$cell.DisplayFormat;$font=$display.Font;$fill=$display.Interior;$borders=$display.Borders
    try{
     $value=$cell.Value2;$has=[bool]$cell.HasFormula;$kind=if($null -eq $value){'Blank'}elseif($value -is [string]){'Text'}elseif($value -is [bool]){'Boolean'}else{'Number'}
     if($has -and [string]$cell.Text -in @('#REF!','#VALUE!','#DIV/0!','#N/A','#NAME?','#NUM!','#NULL!')){$kind='Error';$value=[string]$cell.Text}
     $edges=@();foreach($edge in @(7,10,8,9)){$b=$borders.Item($edge);try{$edges+=@{Line=$b.LineStyle;Weight=$b.Weight;Color=Color $b.Color}}finally{Release-Com $b}}
     $cells+=@{Address=$cell.Address($false,$false);Value=$value;Kind=$kind;Text=[string]$cell.Text;Bold=[bool]$font.Bold;Italic=[bool]$font.Italic;Strike=[bool]$font.Strikethrough;
      Underline=$font.Underline;FontColor=Color $font.Color;FillColor=if($fill.Pattern -eq -4142){$null}else{Color $fill.Color};NumberFormat=$display.NumberFormat;Borders=$edges}
    }finally{Release-Com $borders;Release-Com $fill;Release-Com $font;Release-Com $display;Release-Com $cell}
   }}
   $result+=@{Sheet=$s.Name;Rules=$formats;Cells=$cells}
  }finally{Release-Com $rules;Release-Com $used;Release-Com $s}
 };return ,$result
}
function Snapshot([string]$name){
 $excel.CalculateFullRebuild();$before=Capture $script:book;$path=Join-Path $target ($name+'.xlsx')
 $script:book.SaveAs($path,51);$script:book.Close($false);Release-Com $script:book;$script:book=$null
 $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;$script:book=$excel.Workbooks.Open($path,0,$false);$excel.CalculateFullRebuild();$after=Capture $script:book
 $stable=(($before|ConvertTo-Json -Depth 15 -Compress)-eq($after|ConvertTo-Json -Depth 15 -Compress))
 @{Path=$path;Sha256=$hash;Before=$before;After=$after;Stable=$stable}|ConvertTo-Json -Depth 15|Set-Content (Join-Path $target ($name+'.json')) -Encoding utf8
 if(-not $stable){throw "Conditional formatting changed after save/reopen: $name"}
}
$excel=$null;$script:book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('ConditionalExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ConditionalExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
 [uint32]$ownedExcelProcess=0;[void][ConditionalExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
 if($WorkbookPath){
  $source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$script:book=$excel.Workbooks.Open($source,0,$true)
  Snapshot 'output';@{Sha256=$hash;Unchanged=$hash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content (Join-Path $target 'source.json');return
 }
 $script:book=$excel.Workbooks.Add(-4167);$s=$book.Worksheets.Item(1);$s.Name='Comparison'
 $values=@($null,-1,0,1,3,5,7,2.5,'3','red',$true,$false,'','😀')
 try{
  for($op=1;$op -le 8;$op++){Values $s $op $values;$letter=[string][char](64+$op);$rule=CellRule $s ($letter+'2:'+ $letter+'15') $op '1';try{Paint $rule 13551615 255 $true}finally{Release-Com $rule}}
 }finally{Release-Com $s}
 $s=Sheet 'Text'
 try{
  for($op=0;$op -le 3;$op++){
   Values $s ($op+1) @($null,'ab','AB','xaby','abc','zab','a?b','123',123,$true,'',' aB ','*ab*')
   $letter=[string][char](65+$op);$range=$s.Range($letter+'2:'+ $letter+'14');$rules=$range.FormatConditions
   try{$rule=$rules.Add(9,[Type]::Missing,[Type]::Missing,[Type]::Missing,'ab',$op);try{Paint $rule 13561798 32768}finally{Release-Com $rule}}
   finally{Release-Com $rules;Release-Com $range}
  }
 }finally{Release-Com $s}
 $s=Sheet 'Duplicates'
 try{
  for($col=1;$col -le 2;$col++){
   Values $s $col @($null,'red','RED',3,'3',3,$true,$true,$false,'','001','1','😀','😀','TRUE','FALSE',' 3 ','3.0','3e0','100%','*','?','a','ab','~*','1,000',1000,'é','É','０１','01','0.5','50%',0.5)
   $letter=[string][char](64+$col);$range=$s.Range($letter+'2:'+ $letter+'35');$rules=$range.FormatConditions
   try{$rule=$rules.AddUniqueValues();try{$rule.DupeUnique=$col-1;Paint $rule 65535}finally{Release-Com $rule}}finally{Release-Com $rules;Release-Com $range}
  }
 }finally{Release-Com $s}
 $s=Sheet 'Priority'
 try{
  foreach($col in 1..2){
   Values $s $col @(0,1,2,3,4,5,6);$letter=[string][char](64+$col);$address=$letter+'2:'+ $letter+'8'
   $base=$s.Range($address);$font=$base.Font;try{$font.Bold=$true}finally{Release-Com $font;Release-Com $base}
   $a=CellRule $s $address 5 '1';$b=CellRule $s $address 5 '3';$c=Expression $s $address '=TRUE'
   try{
    Paint $a 13421823 255 $false;Paint $b 65535 16711680 $true;Paint $c 13434828 0 $true
    $a.SetFirstPriority();$b.Priority=$a.Priority+1;$c.Priority=$b.Priority+1;$a.StopIfTrue=($col -eq 2)
    $font=$b.Font;try{$font.Italic=$true}finally{Release-Com $font}
    $font=$c.Font;try{$font.Underline=2;$font.Strikethrough=$true}finally{Release-Com $font}
    $borders=$c.Borders;$edge=$borders.Item(-4107);try{$edge.LineStyle=1;$edge.Weight=2;$edge.Color=0}finally{Release-Com $edge;Release-Com $borders}
   }finally{Release-Com $c;Release-Com $b;Release-Com $a}
  }
 }finally{Release-Com $s}
 $s=Sheet 'Structure'
 try{
  Values $s 1 @(0,1,2,3,4,5,6);Values $s 2 @(10,11,12,13,14,15,16)
  $rule=Expression $s 'B2:B8' '=$A2>2';try{Paint $rule 13434828}finally{Release-Com $rule}
 }finally{Release-Com $s}
 $s=Sheet 'Numbers'
 try{
  Values $s 1 @(0,1.25,-2,10);Values $s 2 @(0,1.25,-2,10)
  $rule=CellRule $s 'A2:A5' 5 '0';try{$rule.NumberFormat='0.00" units"';Paint $rule 65535}finally{Release-Com $rule}
  $rule=Expression $s 'B2:B5' '=A2>0';try{$rule.NumberFormat='0.0%';Paint $rule 13434828}finally{Release-Com $rule}
  $wide=$s.Columns.Item(1);try{$wide.ColumnWidth=20}finally{Release-Com $wide}
 }finally{Release-Com $s}
 $s=$book.Worksheets.Item('Numbers')
 try{
  $header=$s.Range('A1');try{Set-Value $header 'Amount'}finally{Release-Com $header}
  $range=$s.Range('A1:A5');$filterChecks=@()
  try{foreach($criterion in @('1.25 units','1.25')) {
   [void]$range.AutoFilter(1,$criterion);$visible=@()
   foreach($r in 2..5){$row=$s.Rows.Item($r);try{if(-not $row.Hidden){$visible+=$r}}finally{Release-Com $row}}
   $filterChecks+=@{Criterion=$criterion;VisibleRows=$visible};$s.AutoFilterMode=$false
  }}finally{Release-Com $range}
  $filterChecks|ConvertTo-Json -Depth 5|Set-Content (Join-Path $target 'conditional-number-filter.json') -Encoding utf8
 }finally{Release-Com $s}
 Snapshot 'initial'
 $s=$book.Worksheets.Item('Structure');$cell=$s.Range('A3');try{Set-Value $cell 9}finally{Release-Com $cell;Release-Com $s};Snapshot 'edited'
 $s=$book.Worksheets.Item('Structure');$row=$s.Rows.Item(4);try{[void]$row.Insert()}finally{Release-Com $row;Release-Com $s};Snapshot 'insert-row'
 $s=$book.Worksheets.Item('Structure');$row=$s.Rows.Item(2);try{[void]$row.Delete()}finally{Release-Com $row;Release-Com $s};Snapshot 'delete-first'
 $s=$book.Worksheets.Item('Structure');$last=$book.Worksheets.Item($book.Worksheets.Count)
 try{$s.Copy([Type]::Missing,$last);$copy=$book.Worksheets.Item($book.Worksheets.Count);try{$copy.Name='Copy'}finally{Release-Com $copy}}finally{Release-Com $last;Release-Com $s};Snapshot 'copy-sheet'
}finally{
 if($script:book){$script:book.Close($false);Release-Com $script:book;$script:book=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
