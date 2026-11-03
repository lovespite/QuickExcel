param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$root = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) {
    throw 'Use a new directory under this repository artifacts.'
}
[void](New-Item -ItemType Directory -Path $target)
$inputs = [Collections.Generic.List[object]]::new()
function Input-Cell($address, $kind, $value) {
    $inputs.Add([ordered]@{ Sheet = 'Data'; Address = $address; Kind = $kind; Value = $value })
}
Input-Cell 'A1' 'Text' 'b'; Input-Cell 'B1' 'Number' 2
Input-Cell 'A2' 'Text' 'A'; Input-Cell 'B2' 'Number' 1
Input-Cell 'A3' 'Text' 'a'; Input-Cell 'B3' 'Number' 1
Input-Cell 'A4' 'Text' 'b'; Input-Cell 'B4' 'Number' 3
Input-Cell 'D1' 'Number' 3; Input-Cell 'D2' 'Text' '2'; Input-Cell 'D3' 'Boolean' $true
Input-Cell 'D4' 'Formula' '=""'; Input-Cell 'D6' 'Number' 1
Input-Cell 'F1' 'Number' 0; Input-Cell 'F3' 'Formula' '=""'
Input-Cell 'F4' 'Text' '0'; Input-Cell 'F5' 'Boolean' $false
$formulas = @(
    '=SEQUENCE(3,2,10,-2)', '=SEQUENCE(3,2)*2+1', '=SEQUENCE(0)', '=SEQUENCE(2,0)',
    '=SEQUENCE(2.9,2.8)', '=SEQUENCE(2,2,0,0)', '=SEQUENCE(3)/0', '=SUM(SEQUENCE(2,3))',
    '=FILTER(SEQUENCE(3,2),SEQUENCE(3)>1)', '=FILTER(SEQUENCE(3,2),SEQUENCE(1,2)=2)',
    '=FILTER(SEQUENCE(3),FALSE)', '=FILTER(SEQUENCE(3),SEQUENCE(3)>10,"empty")',
    '=FILTER(SEQUENCE(3,2),SEQUENCE(2))', '=FILTER(SEQUENCE(3),SEQUENCE(3)/0)',
    '=FILTER(Data!A1:B4,Data!B1:B4>1)', '=SORT(Data!A1:B4)', '=SORT(Data!A1:B4,2,-1)',
    '=UNIQUE(Data!A1:B4)', '=UNIQUE(Data!A1:B4,FALSE,TRUE)', '=SORT(SEQUENCE(2,3),1,-1,TRUE)',
    '=UNIQUE(SEQUENCE(2,3,0,0),TRUE)', '=UNIQUE(SEQUENCE(2,3,0,0),TRUE,TRUE)',
    '=SORT(UNIQUE(Data!A1:B4),2,-1)', '=SORT(Data!D1:D6)', '=SORT(Data!D1:D6,1,-1)',
    '=UNIQUE(Data!F1:F5)', '=UNIQUE(Data!F1:F5,FALSE,TRUE)', '=FILTER(Data!D1:D6,SEQUENCE(6)>0)',
    '=SORT(SEQUENCE(3),0)', '=SORT(SEQUENCE(3),1,0)', '=SORT(SEQUENCE(3),1,2)',
    '=SEQUENCE(3)+SEQUENCE(1,2)', '=SEQUENCE(3)+SEQUENCE(2)', '=SEQUENCE(2,2)+SEQUENCE(3)',
    '=COUNT(SEQUENCE(2,3))', '=AVERAGE(FILTER(SEQUENCE(4),SEQUENCE(4)>2))',
    '=XLOOKUP("b",Data!A1:A4,Data!A1:B4)', '=XLOOKUP("b",Data!A1:B1,Data!A1:B4)',
    '=XLOOKUP("missing",Data!A1:A4,Data!A1:B4,Data!A1:B1)', '=XLOOKUP("b",Data!A1:A4,Data!A1:B4,,0,-1)',
    '=IF(SEQUENCE(3)>1,SEQUENCE(3),9)', '=IF(FALSE,SEQUENCE(3),7)', '=IF(TRUE,SEQUENCE(3),1/0)',
    '=IFERROR(SEQUENCE(3)/0,SEQUENCE(2))', '=IFERROR(SEQUENCE(2)/0,SEQUENCE(3))',
    '=IFNA(SEQUENCE(3)+SEQUENCE(2),9)', '=FILTER(SEQUENCE(3),SEQUENCE(3)>10,SEQUENCE(2,2))',
    '=IFERROR(SEQUENCE(2),SEQUENCE(3))', '=IFNA(SEQUENCE(2),SEQUENCE(3))'
)
$cases = [Collections.Generic.List[object]]::new()
$arrays = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $formulas.Count; $i++) {
    $top = 1 + $i * 10
    $arrays.Add([ordered]@{ Id = "array-$i"; Formula = $formulas[$i]; Row = $top - 1; Column = 0; Rows = 6; Columns = 4 })
    for ($r = 0; $r -lt 6; $r++) {
        for ($c = 0; $c -lt 4; $c++) {
            $address = ([string][char](65 + $c)) + ($top + $r)
            $entry = [ordered]@{ Id = "array-$i-$r-$c"; Sheet = 'Results'; Address = $address }
            if ($r -eq 0 -and $c -eq 0) { $entry.Formula = $formulas[$i] }
            $cases.Add($entry)
        }
    }
}
$spec = [ordered]@{ Sheets = @('Data','Results'); Date1904 = $false; Inputs = $inputs.ToArray(); Cases = $cases.ToArray(); Arrays = $arrays.ToArray(); Edits = @(); GuardedCases = @() }
$specPath = Join-Path $target 'dynamic-array-spec.json'
$spec | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $specPath -Encoding utf8
& (Join-Path $PSScriptRoot 'test-excel-scalar-functions.ps1') -SpecPath $specPath -OutputDirectory (Join-Path $target 'excel')
if (-not $?) { throw 'Excel dynamic array oracle generation failed.' }
$arity = [Collections.Generic.List[object]]::new()
foreach ($entry in @(@('SEQUENCE',1,4), @('FILTER',2,3), @('SORT',1,4), @('UNIQUE',1,3))) {
    foreach ($count in 0..5) {
        $arguments = (@('1') * $count) -join ','
        $arity.Add([ordered]@{ Function = $entry[0]; Count = $count; Formula = ('=' + $entry[0] + '(' + $arguments + ')');
            ExpectedAccepted = ($count -ge $entry[1] -and $count -le $entry[2]) })
    }
}
$arityPath = Join-Path $target 'array-arity-spec.json'
@{ Cases = $arity.ToArray() } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $arityPath -Encoding utf8
& (Join-Path $PSScriptRoot 'test-excel-formula-arity.ps1') -SpecPath $arityPath -OutputDirectory (Join-Path $target 'arity')
if (-not $?) { throw 'Excel array argument boundary comparison failed.' }
Write-Output "Dynamic array source and independent Excel oracle: $target"
