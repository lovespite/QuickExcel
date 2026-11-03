param([Parameter(Mandatory)][string]$SpecPath, [Parameter(Mandatory)][string]$OutputDirectory,
    [string]$WorkbookPath, [string]$OraclePath, [ValidateSet('Initial','Edited')][string]$State='Initial',
    [ValidateSet(51,61)][int]$FileFormat=51)
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
        if($item.NumberFormat){$cell.NumberFormat=[string]$item.NumberFormat}
    }catch{throw "Cannot set $($item.Sheet)!$($item.Address) [$($item.Value)]: $_"}finally{Release-Com $cell;Release-Com $sheet}
}
function Capture($book,$caseList=$spec.Cases){
    $records=[Collections.Generic.List[object]]::new()
    foreach($case in $caseList){
        $sheet=$book.Worksheets.Item($case.Sheet);$cell=$sheet.Range($case.Address)
        try{
            $value=$cell.Value2;$kind='Number';$text=[string]$cell.Text
            if($null -eq $value){$kind='Blank'}elseif($value -is [string]){$kind='Text'}elseif($value -is [bool]){$kind='Boolean'}elseif($text -match '^#(NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!)$'){$kind='Error';$value=$text}
            $records.Add([ordered]@{Id=$case.Id;Sheet=$case.Sheet;Address=$case.Address;Formula=[string]$cell.Formula;Formula2=[string]$cell.Formula2;Kind=$kind;Value=$value})
        }finally{Release-Com $cell;Release-Com $sheet}
    }
    return ,$records.ToArray()
}
$excel=$null;$book=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if(-not ('ScalarExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ScalarExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
    [uint32]$ownedExcelProcess=0;[void][ScalarExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
    $ownedProcess=Get-Process -Id $ownedExcelProcess
    [ordered]@{ProcessId=$ownedExcelProcess;Started=$ownedProcess.StartTime;SpecSha256=(Get-FileHash -LiteralPath $specFile -Algorithm SHA256).Hash;Owner='test-excel-scalar-functions.ps1'}|
        ConvertTo-Json|Set-Content -LiteralPath (Join-Path $target 'excel-owner.json') -Encoding utf8
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
        if($spec.TextCompatibilityVersion){
            if($spec.TextCompatibilityVersion -notin @(1,2)){throw 'Only Excel text compatibility versions 1/2 are acceptance inputs.'}
            $bootstrap=Join-Path $target 'bootstrap.xlsx';$book.SaveAs($bootstrap,51);$book.Close($false);Release-Com $book;$book=$null
            $zip=[IO.Compression.ZipFile]::Open($bootstrap,[IO.Compression.ZipArchiveMode]::Update)
            try{
                $entry=$zip.GetEntry('xl/workbook.xml');$stream=$entry.Open();try{$xml=[Xml.Linq.XDocument]::Load($stream)}finally{$stream.Dispose()}
                $main=[Xml.Linq.XNamespace]'http://schemas.openxmlformats.org/spreadsheetml/2006/main';$compat=[Xml.Linq.XNamespace]'http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion'
                $list=$xml.Root.Element($main+'extLst');if($null -eq $list){$list=[Xml.Linq.XElement]::new($main+'extLst');$xml.Root.Add($list)}
                $extension=@($list.Elements($main+'ext')|Where-Object{$_.Attribute('uri').Value -eq '{D14903EA-33C4-47F7-8F05-3474C54BE107}'})
                if($extension.Count -eq 0){$ext=[Xml.Linq.XElement]::new($main+'ext');$ext.SetAttributeValue('uri','{D14903EA-33C4-47F7-8F05-3474C54BE107}');$list.Add($ext)}else{$ext=$extension[0];$ext.RemoveNodes()}
                $setting=[Xml.Linq.XElement]::new($compat+'version');$setting.SetAttributeValue('setVersion',[int]$spec.TextCompatibilityVersion);$setting.SetAttributeValue('warnBelowVersion',[int]$spec.TextCompatibilityVersion);$ext.Add($setting)
                $entry.Delete();$stream=$zip.CreateEntry('xl/workbook.xml').Open();try{$xml.Save($stream)}finally{$stream.Dispose()}
            }finally{$zip.Dispose()}
            $book=$excel.Workbooks.Open($bootstrap,0,$false)
        }
        foreach($item in $spec.Inputs){Put $book $item}
        foreach($case in $spec.Cases){if($case.Formula){Put $book @{Sheet=$case.Sheet;Address=$case.Address;Kind=$(if($case.Legacy){'FormulaLegacy'}else{'Formula'});Value=$case.Formula}}}
    }
    $excel.CalculateFullRebuild();$initial=Capture $book
    $initialPath=Join-Path $target 'initial.xlsx';$book.SaveAs($initialPath,$FileFormat);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($initialPath,0,$true);$excel.CalculateFullRebuild();$initialReopened=Capture $book
    $stable=($initial|ConvertTo-Json -Depth 5 -Compress) -ceq ($initialReopened|ConvertTo-Json -Depth 5 -Compress)
    $edited=$null;$editedStable=$null
    if(-not $WorkbookPath){
        foreach($item in $spec.Edits){Put $book $item}
        $excel.CalculateFullRebuild();$edited=Capture $book
        $editedPath=Join-Path $target 'edited.xlsx';$book.SaveAs($editedPath,$FileFormat);$book.Close($false);Release-Com $book;$book=$null
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
            $book.SaveAs($guardPath,$FileFormat);$book.Close($false);Release-Com $book;$book=$null
            $guarded+=@{Id=$case.Id;File=[IO.Path]::GetFileName($guardPath);Sha256=(Get-FileHash -LiteralPath $guardPath -Algorithm SHA256).Hash;Observed=$observed}
        }
    }
    $unchanged=if($WorkbookPath){$sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}else{$true}
    $matches=if($WorkbookPath){($oracle.$State|ConvertTo-Json -Depth 5 -Compress) -ceq ($initial|ConvertTo-Json -Depth 5 -Compress)}else{$null}
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=[int]$excel.Calculation;FileFormat=$FileFormat;
        SpecSha256=(Get-FileHash -LiteralPath $specFile -Algorithm SHA256).Hash;Initial=$initial;Edited=$edited;Guarded=$guarded;Stable=$stable;EditedStable=$editedStable;
        SourceUnchanged=$unchanged;SourceSha256=$sourceHash;ReferenceMatches=$matches;InitialSha256=(Get-FileHash -LiteralPath $initialPath -Algorithm SHA256).Hash}|
        ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $target 'scalar-functions.json') -Encoding utf8
    if(-not $stable -or $editedStable -eq $false -or -not $unchanged -or $matches -eq $false){throw 'Excel scalar formula comparison failed; see JSON.'}
    Write-Output "PASS: $($spec.Cases.Count) scalar formulas; full calculation and save/reopen stable."
}finally{if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
