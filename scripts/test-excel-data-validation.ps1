param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath,[switch]$Date1904)
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
function Capture($workbook){
 $records=@()
 for($si=1;$si -le $workbook.Worksheets.Count;$si++){
  $sheet=$workbook.Worksheets.Item($si);$used=$sheet.UsedRange
  try{
   for($r=1;$r -le [Math]::Min(1000,$used.Row+$used.Rows.Count);$r++){
    $cell=$sheet.Cells.Item($r,1);$label=$sheet.Cells.Item($r,2);$validation=$cell.Validation
    try{
     if(-not $label.Value2){continue}
     $value=$cell.Value2;$kind=if($null -eq $value){'Blank'}elseif($value -is [string]){'Text'}elseif($value -is [bool]){'Boolean'}else{'Number'}
     $records+=@{Sheet=$sheet.Name;Address=$cell.Address($false,$false);Case=$label.Value2;Value=$value;Kind=$kind;Valid=[bool]$validation.Value;
      Type=$validation.Type;Operator=$validation.Operator;Formula1=$validation.Formula1;Formula2=$validation.Formula2;Alert=$validation.AlertStyle;
      AllowBlank=[bool]$validation.IgnoreBlank;ShowDropDown=[bool]$validation.InCellDropdown;ShowError=[bool]$validation.ShowError;ShowInput=[bool]$validation.ShowInput;
      InputTitle=$validation.InputTitle;InputMessage=$validation.InputMessage;ErrorTitle=$validation.ErrorTitle;ErrorMessage=$validation.ErrorMessage}
    }finally{Release-Com $validation;Release-Com $label;Release-Com $cell}
   }
  }finally{Release-Com $used;Release-Com $sheet}
 };return ,$records
}
$excel=$null;$book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('ValidationExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ValidationExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
 [uint32]$ownedExcelProcess=0;[void][ValidationExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
 if($WorkbookPath){$source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true)}
 else{
  $book=$excel.Workbooks.Add(-4167);$book.Date1904=[bool]$Date1904;$sheet=$book.Worksheets.Item(1);$sheet.Name='Validation';$script:row=0
  try{
   foreach($item in @(@('D1','red'),@('D2','Green'),@('D3',[double]10))){$cell=$sheet.Range($item[0]);try{Set-Value $cell $item[1]}finally{Release-Com $cell}}
   $names=$book.Names
   try{
    $name=$names.Add('Choices','=Validation!$D$1:$D$3');Release-Com $name
    $name=$names.Add('ChoicesWithBlank','=Validation!$D$1:$D$4');Release-Com $name
   }finally{Release-Com $names}
   function Case([string]$name,[int]$type,[int]$op,$first,$second,$value,[bool]$blank=$true,[int]$alert=1){
    $script:row++;$cell=$sheet.Cells.Item($script:row,1);$label=$sheet.Cells.Item($script:row,2);$v=$cell.Validation
    try{
     $label.Value2=$name
     $f=[string]$first;$f=$f.Replace('{row}',[string]$script:row)
     if($type -in @(3,7)){$v.Add($type,$alert,$op,$f)}else{$v.Add($type,$alert,$op,$f,[string]$second)}
     $v.IgnoreBlank=$blank;$v.InCellDropdown=$true;$v.ShowInput=$true;$v.ShowError=$true
     $v.InputTitle='录入提示';$v.InputMessage='按规则输入';$v.ErrorTitle='验证失败';$v.ErrorMessage='请检查输入值'
     Set-Value $cell $value
    }finally{Release-Com $v;Release-Com $label;Release-Com $cell}
   }
   $values=@($null,-1,0,1,2,3,5,6,2.5,'2','bad',$true,$false,'','😀')
   foreach($type in @(1,2,4,6)){
    foreach($op in 1..8){
     foreach($value in $values){Case ("type"+$type+"/op"+$op+"/"+$script:row) $type $op '1' '5' $value}
    }
   }
   foreach($blank in @($true,$false)){
    foreach($value in @($null,0,1,2,5,6,'2','red','RED','green','10',10,'missing','')){
     Case ("list/"+$blank+"/"+$script:row) 3 1 'red,Green,10' '' $value $blank
     Case ("range/"+$blank+"/"+$script:row) 3 1 '=$D$1:$D$3' '' $value $blank
     Case ("range-blank/"+$blank+"/"+$script:row) 3 1 '=$D$1:$D$4' '' $value $blank
      Case ("named-list/"+$blank+"/"+$script:row) 3 1 '=Choices' '' $value $blank
      Case ("named-list-blank/"+$blank+"/"+$script:row) 3 1 '=ChoicesWithBlank' '' $value $blank
    }
    foreach($value in $values){Case ("custom/"+$blank+"/"+$script:row) 7 1 '=AND(ISNUMBER(A{row}),MOD(A{row},2)=0)' '' $value $blank}
   }
   foreach($value in @($null,0,1,2,5,6)){
    Case ("blank-bound/"+$script:row) 1 1 '=$D$4' '5' $value $true
   }
   Case 'warning' 1 1 '1' '5' 9 $true 2
   Case 'information' 1 1 '1' '5' 9 $true 3
   $shift=if($Date1904){1462}else{0}
   foreach($value in @(46295,46296,46310,46326,46327)){Case ("date-real/"+$script:row) 4 1 '=DATE(2026,10,1)' '=DATE(2026,10,31)' ([double]($value-$shift))}
   foreach($blank in @($true,$false)){foreach($value in @(0,3,9)){
    Case ("custom-blank/"+$blank+"/"+$script:row) 7 1 '=AND(A{row}<5,$D$4>0)' '' $value $blank
    Case ("bound-expression-blank/"+$blank+"/"+$script:row) 1 1 '=$D$4+1' '5' $value $blank
    Case ("custom-range-blank/"+$blank+"/"+$script:row) 7 1 '=COUNTIF($D$1:$D$4,A{row})>0' '' $value $blank
   }}
   $start=$script:row+1;$ranges="A$start"+":A"+($start+1)+",A"+($start+4)+":A"+($start+5);$cells=$sheet.Range($ranges);$v=$cells.Validation
   try{$v.Add(7,1,1,("=A"+$start+"<C"+$start));$v.IgnoreBlank=$true;$v.ShowError=$true;$v.ShowInput=$true}
   finally{Release-Com $v;Release-Com $cells}
   foreach($offset in @(0,1,4,5)){
    $r=$start+$offset;$cell=$sheet.Cells.Item($r,1);$limit=$sheet.Cells.Item($r,3);$label=$sheet.Cells.Item($r,2)
    try{Set-Value $cell ([double]$offset);Set-Value $limit ([double]3);$label.Value2="relative-multi/$r"}
    finally{Release-Com $label;Release-Com $limit;Release-Com $cell}
   }
  }finally{Release-Com $sheet}
 }
 $excel.CalculateFullRebuild();$before=Capture $book;$path=Join-Path $target 'validation.xlsx'
 $book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null;$hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
 $book=$excel.Workbooks.Open($path,0,$true);$excel.CalculateFullRebuild();$after=Capture $book
 $stable=(($before|ConvertTo-Json -Depth 12 -Compress)-eq($after|ConvertTo-Json -Depth 12 -Compress))
 @{Path=$path;Sha256=$hash;Before=$before;After=$after;Stable=$stable}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $target 'validation.json') -Encoding utf8
 if($WorkbookPath){@{Sha256=$sourceHash;Unchanged=$sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content (Join-Path $target 'source.json')}
 if(-not $stable){throw 'Validation changed after save/reopen.'}
}finally{
 if($book){$book.Close($false);Release-Com $book;$book=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($excel){$excel.Quit();Release-Com $excel}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
