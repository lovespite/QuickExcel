param([Parameter(Mandatory)][string]$OutputDirectory,[string]$InputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$output=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$prefix=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $output.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $output)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $output)
function Release-Com($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
$excel=$null;$book=$null;$sheet=$null;$cases=@();$books=@();$today=[datetime]::Today
try {
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 foreach($system in @(1900,1904)) {
  $book=$excel.Workbooks.Add(-4167);$book.Date1904=($system -eq 1904);$sheet=$book.Worksheets.Item(1)
  $serial=$today.ToOADate()-$(if($system -eq 1904){1462}else{0})
  for($col=1;$col -le 10;$col++){
   for($r=1;$r -le 91;$r++){$sheet.Cells.Item($r,$col).Value2=$serial+$r-46+0.5}
   $sheet.Cells.Item(92,$col).NumberFormat='@';$sheet.Cells.Item(92,$col).Formula="'$serial"
   $sheet.Cells.Item(93,$col).NumberFormat='@';$sheet.Cells.Item(93,$col).Formula="'"+$today.ToString('yyyy-MM-dd')
   $sheet.Cells.Item(95,$col).Value2=$true;$sheet.Cells.Item(96,$col).Formula='=1/0'
   $range=$sheet.Range($sheet.Cells.Item(1,$col),$sheet.Cells.Item(96,$col));$rule=$range.FormatConditions.Add(11,[Type]::Missing,[Type]::Missing,[Type]::Missing,[Type]::Missing,[Type]::Missing,$col-1,[Type]::Missing)
   $rule.Interior.Color=255;$excel.CalculateFull()
   $matches=@(1..96 | Where-Object {$cell=$sheet.Cells.Item($_,$col);try{$cell.DisplayFormat.Interior.Color -eq 255}finally{Release-Com $cell}})
   $cases += [ordered]@{system=$system;period=$col-1;matches=$matches;formula=$rule.Formula1}
   Release-Com $rule;Release-Com $range
  }
  $path=Join-Path $output "excel-$system.xlsx";$book.SaveAs($path,51);$book.Close($false);Release-Com $sheet;$sheet=$null;Release-Com $book;$book=$null
  $zip=[IO.Compression.ZipFile]::OpenRead($path);try{$parts=[ordered]@{};foreach($entry in $zip.Entries){if($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')){$reader=[IO.StreamReader]::new($entry.Open());try{$parts[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}}}finally{$zip.Dispose()}
  $books += [ordered]@{system=$system;parts=$parts;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
  if($InputDirectory){
   $source=Join-Path ([IO.Path]::GetFullPath($InputDirectory,$repo)) "$system.xlsx";$book=$excel.Workbooks.Open($source,0,$true);$sheet=$book.Worksheets.Item(1);$excel.CalculateFull()
   foreach($case in $cases|Where-Object system -eq $system){$col=$case.period+1;$matches=@(1..96|Where-Object {$cell=$sheet.Cells.Item($_,$col);try{$cell.DisplayFormat.Interior.Color -eq 255}finally{Release-Com $cell}});if(($matches -join ',') -ne ($case.matches -join ',')){throw "Reopened $system/$($case.period) differs"}}
   $book.Close($false);Release-Com $sheet;$sheet=$null;Release-Com $book;$book=$null
  }
 }
 [ordered]@{today=$today.ToString('yyyy-MM-dd');excelVersion=$excel.Version;excelBuild=$excel.Build;cases=$cases;books=$books}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $output 'excel-date-conditional.json') -Encoding utf8
 Write-Host "PASS: $($cases.Count) independent date conditional-format cases: $output"
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $excel){$excel.Quit()};Release-Com $sheet;Release-Com $book;Release-Com $excel}
