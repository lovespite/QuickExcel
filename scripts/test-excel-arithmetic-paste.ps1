param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$output=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$prefix=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $output.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $output)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $output)
function Release-Com($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
$excel=$null;$book=$null;$sheet=$null;$cases=@();$row=0
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 $book=$excel.Workbooks.Add(-4167);$sheet=$book.Worksheets.Item(1)
 foreach($operation in @('add','subtract','multiply','divide')) {
 foreach($kind in @('numbers','formula-target','formula-source','blank-source','skip-blank','blank-target')) {
  $row++;$source=$sheet.Cells.Item($row,1);$target=$sheet.Cells.Item($row,3)
  try {
   if($kind -eq 'formula-source'){$source.Formula='=1+1'}elseif($kind -notin @('blank-source','skip-blank')){$source.Value2=2}
   if($kind -eq 'formula-target'){$target.Formula='=4+4'}elseif($kind -ne 'blank-target'){$target.Value2=8}
   $op=switch($operation){add{2};subtract{3};multiply{4};divide{5}}
   $null=$source.Copy();$null=$target.PasteSpecial(-4163,$op,($kind -eq 'skip-blank'),$false)
   $excel.CalculateFull()
   $cases += [ordered]@{operation=$operation;kind=$kind;value=$target.Value2;display=$target.Text;formula=$target.Formula;hasFormula=[bool]$target.HasFormula}
  }finally{Release-Com $source;Release-Com $target}
 }
 }
 $path=Join-Path $output 'excel-arithmetic-paste.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $sheet;$sheet=$null;Release-Com $book;$book=$null
 [ordered]@{excelVersion=$excel.Version;excelBuild=$excel.Build;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;cases=$cases}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $output 'excel-arithmetic-paste.json') -Encoding utf8
 Write-Host "PASS: $($cases.Count) arithmetic paste cases captured: $output"
}finally{if($null -ne $book){$book.Close($false)};if($null -ne $excel){$excel.Quit()};Release-Com $sheet;Release-Com $book;Release-Com $excel}
