param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop';$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$excel=$null;$book=$null;$sheet=$null;$records=@()
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 $book=$excel.Workbooks.Add(-4167);$sheet=$book.Worksheets.Item(1)
 $cell=$sheet.Range('A1');try{$cell.Value2='Value'}finally{Release-Com $cell}
 foreach($r in 2..12){$cell=$sheet.Cells.Item($r,1);try{$cell.Value2=[double]($r-1)}finally{Release-Com $cell}}
 $cell=$sheet.Range('B1');try{$cell.Formula='=SUBTOTAL(9,A2:A12)'}finally{Release-Com $cell}
 $cell=$sheet.Range('C1');try{$cell.Formula='=SUBTOTAL(109,A2:A12)'}finally{Release-Com $cell}
 foreach($range in @('3:5','2:12')){$row=$sheet.Rows.Item($range);try{[void]$row.Group()}finally{Release-Com $row}}
 foreach($stage in @('grouped','filtered','child-collapsed','parent-collapsed','parent-expanded','filter-cleared')){
  if($stage -eq 'filtered'){$range=$sheet.Range('A1:A12');try{[void]$range.AutoFilter(1,'>=6')}finally{Release-Com $range}}
  if($stage -eq 'child-collapsed'){$row=$sheet.Rows.Item(6);try{$row.ShowDetail=$false}finally{Release-Com $row}}
  if($stage -in @('parent-collapsed','parent-expanded')){$row=$sheet.Rows.Item(13);try{$row.ShowDetail=($stage -eq 'parent-expanded')}finally{Release-Com $row}}
  if($stage -eq 'filter-cleared'){[void]$sheet.ShowAllData()}
  $excel.CalculateFullRebuild();$hidden=@();foreach($r in 1..14){$row=$sheet.Rows.Item($r);try{if($row.Hidden){$hidden+=$r}}finally{Release-Com $row}}
  $r6=$sheet.Rows.Item(6);$r13=$sheet.Rows.Item(13);$b=$sheet.Range('B1');$c=$sheet.Range('C1')
  try{$records+=@{Stage=$stage;Hidden=$hidden;ChildDetail=[bool]$r6.ShowDetail;ParentDetail=[bool]$r13.ShowDetail;Subtotal9=$b.Value2;Subtotal109=$c.Value2}}finally{Release-Com $r6;Release-Com $r13;Release-Com $b;Release-Com $c}
  $book.SaveAs((Join-Path $target ($stage+'.xlsx')),51)
 }
 Release-Com $sheet;$sheet=$null;$book.Close($false);Release-Com $book;$book=$null
 $hashes=@{};foreach($stage in $records.Stage){$path=Join-Path $target ($stage+'.xlsx');$hashes[$stage]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($path,0,$true);$excel.CalculateFullRebuild();$book.Close($false);Release-Com $book;$book=$null}
 @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);NormalReopen=$true;Hashes=$hashes;Cases=$records}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $target 'outline-filter.json') -Encoding utf8
}finally{
 Release-Com $sheet;if($book){$book.Close($false);Release-Com $book;$book=$null};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}
}
