param([Parameter(Mandatory)][string]$SpecPath, [Parameter(Mandatory)][string]$OutputDirectory,
    [string]$WorkbookPath, [string]$OraclePath, [ValidateSet('Initial','Edited')][string]$State='Initial')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent; $target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
$specFile=[IO.Path]::GetFullPath($SpecPath,$repo); $spec=Get-Content -LiteralPath $specFile -Raw|ConvertFrom-Json
Copy-Item -LiteralPath $specFile -Destination (Join-Path $target 'spec.json')
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Put($book,$item){
    $sheet=$book.Worksheets.Item($item.Sheet);$cell=$sheet.Range($item.Address)
    try{
        switch($item.Kind){
            'Blank'{$cell.ClearContents()}
            'Number'{$cell.Value2=[double]$item.Value}
            'Text'{$cell.NumberFormat='@';$cell.Value2=[string]$item.Value}
            'Boolean'{$cell.Value2=[bool]$item.Value}
            'Formula'{$cell.Formula2=[string]$item.Value}
            'FormulaLegacy'{$cell.Formula=[string]$item.Value}
            default{throw 'Invalid input kind.'}
        }
    }catch{throw "Cannot set $($item.Sheet)!$($item.Address) [$($item.Value)]: $_"}finally{Release-Com $cell;Release-Com $sheet}
}
function Capture($book,$caseList=$spec.Cases){
    $records=[Collections.Generic.List[object]]::new()
    foreach($case in $caseList){
        $sheet=$book.Worksheets.Item($case.Sheet);$cell=$sheet.Range($case.Address)
        try{
            $value=$cell.Value2;$kind='Number';$text=[string]$cell.Text
            if($null -eq $value){$kind='Blank'}elseif($value -is [string]){$kind='Text'}elseif($value -is [bool]){$kind='Boolean'}elseif($text -match '^#'){$kind='Error';$value=$text}
            $records.Add([ordered]@{Id=$case.Id;Sheet=$case.Sheet;Address=$case.Address;Formula=[string]$cell.Formula;Formula2=[string]$cell.Formula2;Kind=$kind;Value=$value})
        }finally{Release-Com $cell;Release-Com $sheet}
    }
    return ,$records.ToArray()
}
$excel=$null;$book=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if($WorkbookPath){
        if(-not $OraclePath){throw 'Verification requires OraclePath.'}
        $source=[IO.Path]::GetFullPath($WorkbookPath,$repo);$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        $oracle=Get-Content -LiteralPath $OraclePath -Raw|ConvertFrom-Json
        $book=$excel.Workbooks.Open($source,0,$true)
    }else{
        $book=$excel.Workbooks.Add();$book.Date1904=[bool]$spec.Date1904
        while($book.Worksheets.Count -gt 1){$extra=$book.Worksheets.Item($book.Worksheets.Count);$extra.Delete();Release-Com $extra}
        $first=$book.Worksheets.Item(1);$first.Name=$spec.Sheets[0];Release-Com $first
        foreach($name in $spec.Sheets|Select-Object -Skip 1){$last=$book.Worksheets.Item($book.Worksheets.Count);try{$sheet=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last};try{$sheet.Name=$name}finally{Release-Com $sheet}}
        foreach($item in $spec.Inputs){Put $book $item}
        foreach($case in $spec.Cases){Put $book @{Sheet=$case.Sheet;Address=$case.Address;Kind=$(if($case.Legacy){'FormulaLegacy'}else{'Formula'});Value=$case.Formula}}
    }
    $excel.CalculateFullRebuild();$initial=Capture $book
    $initialPath=Join-Path $target 'initial.xlsx';$book.SaveAs($initialPath,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($initialPath,0,$true);$excel.CalculateFullRebuild();$initialReopened=Capture $book
    $stable=($initial|ConvertTo-Json -Depth 5 -Compress) -ceq ($initialReopened|ConvertTo-Json -Depth 5 -Compress)
    $edited=$null;$editedStable=$null
    if(-not $WorkbookPath){
        foreach($item in $spec.Edits){Put $book $item}
        $excel.CalculateFullRebuild();$edited=Capture $book
        $editedPath=Join-Path $target 'edited.xlsx';$book.SaveAs($editedPath,51);$book.Close($false);Release-Com $book;$book=$null
        $book=$excel.Workbooks.Open($editedPath,0,$true);$excel.CalculateFullRebuild();$editedReopened=Capture $book
        $editedStable=($edited|ConvertTo-Json -Depth 5 -Compress) -ceq ($editedReopened|ConvertTo-Json -Depth 5 -Compress)
    }
    $book.Close($false);Release-Com $book;$book=$null
    $guarded=@()
    if(-not $WorkbookPath){
        foreach($case in $spec.GuardedCases){
            $book=$excel.Workbooks.Open($initialPath,0,$true)
            Put $book @{Sheet=$case.Sheet;Address=$case.Address;Kind='Formula';Value=$case.Formula}
            $excel.CalculateFullRebuild();$observed=Capture $book @($case);$guardPath=Join-Path $target ('guard-'+$case.Id+'.xlsx')
            $book.SaveAs($guardPath,51);$book.Close($false);Release-Com $book;$book=$null
            $guarded+=@{Id=$case.Id;File=[IO.Path]::GetFileName($guardPath);Sha256=(Get-FileHash -LiteralPath $guardPath -Algorithm SHA256).Hash;Observed=$observed}
        }
    }
    $unchanged=if($WorkbookPath){$sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}else{$true}
    $matches=if($WorkbookPath){($oracle.$State|ConvertTo-Json -Depth 5 -Compress) -ceq ($initial|ConvertTo-Json -Depth 5 -Compress)}else{$null}
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=[int]$excel.Calculation;
        SpecSha256=(Get-FileHash -LiteralPath $specFile -Algorithm SHA256).Hash;Initial=$initial;Edited=$edited;Guarded=$guarded;Stable=$stable;EditedStable=$editedStable;
        SourceUnchanged=$unchanged;SourceSha256=$sourceHash;ReferenceMatches=$matches;InitialSha256=(Get-FileHash -LiteralPath $initialPath -Algorithm SHA256).Hash}|
        ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $target 'scalar-functions.json') -Encoding utf8
    if(-not $stable -or $editedStable -eq $false -or -not $unchanged -or $matches -eq $false){throw 'Excel scalar formula comparison failed; see JSON.'}
    Write-Output "PASS: $($spec.Cases.Count) scalar formulas; full calculation and save/reopen stable."
}finally{if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
