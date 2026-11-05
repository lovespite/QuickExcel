param([Parameter(Mandatory)][string]$OutputDirectory, [string]$InputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$output=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$prefix=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $output.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $output)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $output)
Add-Type -TypeDefinition @'
using System;
using System.Reflection;
public static class ExcelDateGroupOracle {
 public static void Apply(object range, int[] levels, DateTime[] dates) {
  var criteria=new object[levels.Length*2];
  for(int i=0;i<levels.Length;i++){criteria[i*2]=levels[i];criteria[i*2+1]=dates[i];}
  range.GetType().InvokeMember("AutoFilter",BindingFlags.InvokeMethod,null,range,new object[]{1,Type.Missing,7,criteria,Type.Missing,Type.Missing});
 }
}
'@
function Release-Com($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
function Visible-Rows($sheet){@(2..10 | Where-Object {$row=$sheet.Rows.Item($_);try{-not $row.Hidden}finally{Release-Com $row}})}
$excel=$null;$book=$null;$sheet=$null;$cases=@()
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 foreach($system in @(1900,1904)) {
 foreach($kind in @('year','month','day','multiple')) {
  $book=$excel.Workbooks.Add(-4167);$book.Date1904=($system -eq 1904);$sheet=$book.Worksheets.Item(1)
  $sheet.Range('A1').Value2='Date'
  $dates=@([datetime]'2026-11-05T12:30:00',[datetime]'2026-11-30',[datetime]'2026-12-01',[datetime]'2025-11-05',[datetime]'2026-11-05T12:30:00')
  for($i=0;$i -lt $dates.Count;$i++){$row=$i+2;$cell=$sheet.Cells.Item($row,1);try{$cell.Value2=$dates[$i].ToOADate()-$(if($system -eq 1904){1462}else{0})}finally{Release-Com $cell}}
  $sheet.Range('A2:A5').NumberFormat='yyyy-mm-dd';$sheet.Range('A7').NumberFormat='@';$sheet.Range('A7').Value2='2026-11-05'
  $sheet.Range('A9').Value2=$true;$sheet.Range('A10').Formula='=1/0'
  [int[]]$levels=switch($kind){year{0};month{1};day{2};multiple{0;2}}
  [datetime[]]$criteriaDates=if($kind -eq 'multiple'){@([datetime]'2025-01-01',[datetime]'2026-12-01')}else{@([datetime]'2026-11-05')}
  [ExcelDateGroupOracle]::Apply($sheet.Range('A1:A10'),$levels,$criteriaDates)
  $visible=@(Visible-Rows $sheet);$name="$system-$kind";$path=Join-Path $output "excel-$name.xlsx";$book.SaveAs($path,51)
  $book.Close($false);Release-Com $sheet;$sheet=$null;Release-Com $book;$book=$null
  $zip=[IO.Compression.ZipFile]::OpenRead($path)
  try{
   $parts=[ordered]@{}
   foreach($entry in $zip.Entries){if($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')){$reader=[IO.StreamReader]::new($entry.Open());try{$parts[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}
  }finally{$zip.Dispose()}
  if($InputDirectory){
   $source=Join-Path ([IO.Path]::GetFullPath($InputDirectory,$repo)) "$name.xlsx"
   $book=$excel.Workbooks.Open($source,0,$true);$sheet=$book.Worksheets.Item(1);$sheet.AutoFilter.ApplyFilter()
   $actual=@(Visible-Rows $sheet);if(($actual -join ',') -ne ($visible -join ',')){throw "$name reapply differs: $actual"}
   $book.SaveAs((Join-Path $output "roundtrip-$name.xlsx"),51)
   $book.Close($false);Release-Com $sheet;$sheet=$null;Release-Com $book;$book=$null
  }
  $cases += [ordered]@{name=$name;system=$system;kind=$kind;visibleRows=$visible;sourceSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;parts=$parts}
 }
 }
 [ordered]@{excelVersion=$excel.Version;excelBuild=$excel.Build;cases=$cases}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $output 'excel-date-filters.json') -Encoding utf8
 Write-Host "PASS: independent date filter cases captured: $output"
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $excel){$excel.Quit()};Release-Com $sheet;Release-Com $book;Release-Com $excel}
