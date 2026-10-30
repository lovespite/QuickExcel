param([Parameter(Mandatory)][string]$OutputDirectory, [string]$WorkbookPath, [string]$ReferencePath, [double]$ExpectedValue = 9)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$artifactsRoot = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($artifactsRoot, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $target)
$excel = $null; $book = $null; $sheet = $null
if (-not ('QuickExcelPropertyInterop' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Reflection;
public static class QuickExcelPropertyInterop {
    public static object Get(object target, string name, object[] args) => target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, args);
    public static void Set(object target, string name, object value) => target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] {value});
    public static object Add(object target, string name, int type, object value) => target.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, target, new object[] {name, false, type, value});
}
'@
}
function Release-Com($object) { if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) } }
function Get-ComProperty($object, [string]$name, [object[]]$arguments = @()) {
    return ,([QuickExcelPropertyInterop]::Get($object, $name, $arguments))
}
function Set-ComProperty($object, [string]$name, $value) {
    [QuickExcelPropertyInterop]::Set($object, $name, $value)
}
function Add-CustomProperty($object, [string]$name, [int]$type, $value) {
    $property=[QuickExcelPropertyInterop]::Add($object, $name, $type, $value)
    Release-Com $property
}
function Read-Evidence($workbook) {
    $built = $workbook.BuiltinDocumentProperties; $custom = $workbook.CustomDocumentProperties; $worksheet = $workbook.Worksheets.Item(1)
    try {
        $properties = [ordered]@{}
        foreach ($index in @(1,2,3,4,5,20,21)) { $property = Get-ComProperty $built 'Item' @($index); try { $properties[[string]$index] = [string](Get-ComProperty $property 'Value') } finally { Release-Com $property } }
        $customValues = [ordered]@{}
        foreach ($name in @('报价号', '已批准', '费率')) { $property = Get-ComProperty $custom 'Item' @($name); try { $customValues[$name] = Get-ComProperty $property 'Value' } finally { Release-Com $property } }
        $first = $worksheet.Cells.Item(1,1); $formula = $worksheet.Cells.Item(1,2)
        try { return [ordered]@{Properties=$properties; Custom=$customValues; Sheet=[string]$worksheet.Name; Input=$first.Value2; Formula=$formula.Formula; Result=$formula.Value2; Text=$formula.Text} }
        finally { Release-Com $first; Release-Com $formula }
    } finally { Release-Com $built; Release-Com $custom; Release-Com $worksheet }
}
try {
    $sourceHash = if ($WorkbookPath) { (Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash } else { $null }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false; $excel.AutomationSecurity=3
    if ($WorkbookPath) { $book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath,$repo),0,$true) }
    else {
        $book=$excel.Workbooks.Add(); $ExpectedValue=2
        while ($book.Worksheets.Count -gt 1) { $extra=$book.Worksheets.Item($book.Worksheets.Count); try {$extra.Delete()} finally {Release-Com $extra} }
        $sheet=$book.Worksheets.Item(1); $sheet.Name='属性对照'
        $first=$sheet.Cells.Item(1,1); $formula=$sheet.Cells.Item(1,2)
        try {$first.Value2=[double]2; $formula.Formula='=SUM(A1,3)'} finally {Release-Com $first; Release-Com $formula}
        $built=$book.BuiltinDocumentProperties; $custom=$book.CustomDocumentProperties
        try {
            $values=@{1='报价与预算';2='XLSX 保存验收';3='QuickExcel fixture';4='metadata';5='仅供自动验收';20='测试经理';21='测试公司'}
            foreach ($index in $values.Keys) {$property=Get-ComProperty $built 'Item' @([int]$index); try {Set-ComProperty $property 'Value' $values[$index]} finally {Release-Com $property}}
            Add-CustomProperty $custom '报价号' 4 '00123'
            Add-CustomProperty $custom '已批准' 2 $true
            Add-CustomProperty $custom '费率' 5 ([double]0.15)
        } finally {Release-Com $built;Release-Com $custom}
    }
    $excel.CalculateFullRebuild()
    $calculationMode=$excel.Calculation
    $firstRead=Read-Evidence $book
    if ($ReferencePath) {
        $reference=(Get-Content -LiteralPath ([IO.Path]::GetFullPath($ReferencePath,$repo)) -Raw | ConvertFrom-Json).AfterReopen
        if (($reference.Properties | ConvertTo-Json -Compress) -ne ($firstRead.Properties | ConvertTo-Json -Compress) -or
            $reference.Custom.'报价号' -ne $firstRead.Custom['报价号'] -or $reference.Custom.'已批准' -ne $firstRead.Custom['已批准'] -or
            [Math]::Abs($reference.Custom.'费率' - $firstRead.Custom['费率']) -gt 1e-12) {throw 'QuickExcel output did not retain all reference document properties.'}
    }
    if ($firstRead.Input -ne $ExpectedValue -or $firstRead.Result -ne ($ExpectedValue+3) -or $firstRead.Formula -ne '=SUM(A1,3)') {throw 'Excel formula/value comparison failed.'}
    $savedPath=Join-Path $target $(if ($WorkbookPath) {'excel-resaved.xlsx'} else {'excel-source.xlsx'})
    $book.SaveAs($savedPath,51); $book.Close($false); Release-Com $book; $book=$null
    Release-Com $sheet; $sheet=$null
    $book=$excel.Workbooks.Open($savedPath,0,$true); $excel.CalculateFullRebuild(); $secondRead=Read-Evidence $book
    $firstComparison=$firstRead | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $secondComparison=$secondRead | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $rateMatches=[Math]::Abs($firstRead.Custom['费率']-$secondRead.Custom['费率']) -le 1e-12
    $firstComparison.Custom.'费率'=$null; $secondComparison.Custom.'费率'=$null
    $sameAfterReopen=$rateMatches -and ($firstComparison | ConvertTo-Json -Depth 8 -Compress) -eq ($secondComparison | ConvertTo-Json -Depth 8 -Compress)
    $book.Close($false); Release-Com $book; $book=$null
    $sourceUnchanged= -not $WorkbookPath -or $sourceHash -eq (Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash
    [ordered]@{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=$calculationMode;NumericPropertyTolerance=1e-12;
        CorruptLoad='xlNormalLoad'; SourceUnchanged=$sourceUnchanged; SameAfterReopen=$sameAfterReopen; BeforeSave=$firstRead; AfterReopen=$secondRead; SavedSha256=(Get-FileHash -LiteralPath $savedPath).Hash} |
        ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $target 'excel-metadata.json') -Encoding utf8
    if (-not $sourceUnchanged) {throw 'Input file was modified.'}
    if (-not $sameAfterReopen) {throw 'Excel save/reopen changed properties, formula or result. See excel-metadata.json.'}
    Write-Output $target
} finally {
    if ($book) {$book.Close($false); Release-Com $book}
    Release-Com $sheet
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excel) {$excel.Quit(); Release-Com $excel}
}
