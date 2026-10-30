param([Parameter(Mandatory)][string]$OutputDirectory, [string]$WorkbookPath, [string]$OraclePath, [ValidateSet('Initial','Edited')][string]$State='Initial')
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent; $target = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$root = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } }
function Put($sheet, $cellData) {
    $cell = $sheet.Range($cellData.Address)
    try {
        switch ($cellData.Kind) {
            'Number' { $cell.Value2 = [double]$cellData.Value }
            'Text' { $cell.NumberFormat = '@'; $cell.Value2 = [string]$cellData.Value }
            'Boolean' { $cell.Value2 = [bool]$cellData.Value }
            'Formula' { $cell.Formula = [string]$cellData.Value }
        }
    } catch { throw "Cannot set $($cellData.Address) [$($cellData.Value)]: $_" } finally { Release-Com $cell }
}
function Capture($book, $cases) {
    $sheet = $book.Worksheets.Item('Results'); $items = @()
    try { foreach ($case in $cases) {
        $cell = $sheet.Range($case.Address)
        try {
            $value = $cell.Value2; $kind = 'Number'
            if ([string]$cell.Text -match '^#') { $kind = 'Error'; $value = [string]$cell.Text }
            elseif ($null -eq $value) { $kind = 'Blank' } elseif ($value -is [bool]) { $kind = 'Boolean' } elseif ($value -is [string]) { $kind = 'Text' }
            $items += [ordered]@{Id=$case.Id;Address=$case.Address;Formula=[string]$cell.Formula;Kind=$kind;Value=$value}
        } finally { Release-Com $cell }
    } } finally { Release-Com $sheet }
    return ,$items
}
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false; $excel.AutomationSecurity=3
    $inputs = @(); $rejected = @(); $cases = [Collections.Generic.List[object]]::new()
    function Add-Case([string]$id, [string]$formula) { $cases.Add([ordered]@{Id=$id;Address=('A'+($cases.Count+1));Formula=$formula}) }
    if ($WorkbookPath) {
        if (-not $OraclePath) { throw 'Verification requires OraclePath.' }
        $oracle = Get-Content -LiteralPath $OraclePath -Raw | ConvertFrom-Json
        $cases.AddRange([object[]]$oracle.Cases); $inputs = $oracle.Inputs
        $source = [IO.Path]::GetFullPath($WorkbookPath, $repo); $sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        $book = $excel.Workbooks.Open($source, 0, $true)
    } else {
        $values = @(
            @('Blank',$null), @('Formula','=""'), @('Number',0), @('Text','0'), @('Number',1), @('Text','01'), @('Number',-1),
            @('Boolean',$true), @('Boolean',$false), @('Text','TRUE'), @('Text','apple'), @('Text','Apple'), @('Text','a*b'), @('Text','a?b'),
            @('Text','a~b'), @('Text','#N/A'), @('Formula','=#N/A'), @('Formula','=1/0'), @('Text',' '), @('Number',2.5), @('Text','2.5'),
            @('Formula','=#VALUE!'), @('Text','abc'), @('Text','~'), @('Text','.'), @('Text','1/2/2020'), @('Number',43832))
        for ($i=0; $i -lt $values.Count; $i++) {
            $inputs += [ordered]@{Address=('A'+($i+1));Kind=$values[$i][0];Value=$values[$i][1]}
            $kind='Number'; $value=$i+1
            if ($i -eq 1) { $kind='Blank';$value=$null } elseif ($i -eq 3) { $kind='Text';$value='10' }
            elseif ($i -eq 7) { $kind='Boolean';$value=$true } elseif ($i -eq 8) { $kind='Boolean';$value=$false }
            elseif ($i -eq 16) { $kind='Formula';$value='=#REF!' } elseif ($i -eq 21) { $kind='Formula';$value='=#N/A' }
            $inputs += [ordered]@{Address=('B'+($i+1));Kind=$kind;Value=$value}
            $inputs += [ordered]@{Address=('C'+($i+1));Kind='Text';Value=('group'+($i%3))}
        }
        foreach ($item in @(@('D1','Blank',$null),@('D2','Formula','=""'),@('D3','Formula','=#N/A'),@('D4','Boolean',$true),@('D5','Boolean',$false),@('D6','Number',0),@('D7','Text','1'),@('D8','Text',('x'*256)))) {
            $inputs += [ordered]@{Address=$item[0];Kind=$item[1];Value=$item[2]}
        }
        foreach ($item in @(@('E1','Number',1000),@('E2','Number',0.025),@('E3','Text','2,5'),@('E4','Text','1,000'),@('E5','Number',25),@('E6','Text','a~~b'),@('E7','Text','a~b'))) {
            $inputs += [ordered]@{Address=$item[0];Kind=$item[1];Value=$item[2]}
        }
        $criteria = @('""','"="','"<>"','0','"0"','"=0"','"<>0"','">0"','">=0"','"<0"','"<=0"','1','"01"','"=01"','">=1"',
            'TRUE','FALSE','"TRUE"','"=TRUE"','"<>TRUE"','">FALSE"','"*"','"?"','"a*"','"<>a*"','"a~*b"','"a~?b"','"a~~b"','"~"','"~~"','"~a"',
            '"#N/A"','"<>#N/A"','"#DIV/0!"','"#VALUE!"','">apple"','"<=apple"','" "','Data!D1','Data!D2','Data!D3','Data!D4','Data!D5',
            '"2.5"','"2,5"','"2.5%"','"1/2/2020"','"2020/1/2"','Data!D8','"="&Data!D6')
        for ($c=0; $c -lt $criteria.Count; $c++) {
            $criterion=$criteria[$c]
            for ($r=1; $r -le $values.Count; $r++) { Add-Case "match-$c-$r" ('=COUNTIF(Data!A'+$r+','+$criterion+')') }
            foreach ($name in @('COUNTIF','COUNTIFS','SUMIF','SUMIFS','AVERAGEIF','AVERAGEIFS')) {
                $formula = switch ($name) {
                    'COUNTIF' { '=COUNTIF(Data!A1:A27,'+$criterion+')' }
                    'COUNTIFS' { '=COUNTIFS(Data!A1:A27,'+$criterion+')' }
                    'SUMIF' { '=SUMIF(Data!A1:A27,'+$criterion+',Data!B1:B27)' }
                    'SUMIFS' { '=SUMIFS(Data!B1:B27,Data!A1:A27,'+$criterion+')' }
                    'AVERAGEIF' { '=AVERAGEIF(Data!A1:A27,'+$criterion+',Data!B1:B27)' }
                    'AVERAGEIFS' { '=AVERAGEIFS(Data!B1:B27,Data!A1:A27,'+$criterion+')' }
                }
                Add-Case "$name-$c" $formula
            }
        }
        $extra = @(
            '=COUNTBLANK(Data!A1:A27)', '=COUNTBLANK(Data!A1)', '=COUNTBLANK(Data!A2)', '=COUNTBLANK(Data!A:A)', '=COUNTIF(Data!A:A,"")', '=COUNTIFS(Data!A:A,"",Data!B:B,"")',
            '=SUMIF(Data!A1:A5,">=0",Data!B1)', '=AVERAGEIF(Data!A1:A5,">=0",Data!B1)', '=SUMIF(Data!A1:A5,">=0",Data!B1:C10)',
            '=SUMIFS(Data!B1:B5,Data!A1:A4,">0")', '=COUNTIFS(Data!A1:A5,">0",Data!C1:C4,"group1")', '=AVERAGEIFS(Data!B1:B5,Data!A1:A4,">0")',
            '=SUMIF(Data!A1:A27,"<>0")', '=AVERAGEIF(Data!A1:A27,"<>0")', '=SUMIF(Data!A1:A27,"*",Data!B1)',
            '=SUMIFS(Data!B1:B27,Data!A1:A27,">0",Data!C1:C27,"group1")', '=COUNTIFS(Data!A1:A27,"*",Data!C1:C27,"group1")',
            '=AVERAGEIFS(Data!B1:B27,Data!A1:A27,"*",Data!C1:C27,"<>group1")', '=COUNTIF(Data!A1:A27,">"&Data!D6)',
            '=SUMIF(Data!A1:A27,"nomatch",Data!B1:B27)', '=AVERAGEIF(Data!A1:A27,"nomatch",Data!B1:B27)', '=AVERAGEIFS(Data!B1:B27,Data!A1:A27,"nomatch")',
            '=SUMIFS(Data!B1:B27,Data!A1:A27,"#N/A")', '=SUMIF(Data!A1:A27,"#VALUE!",Data!B1)',
            '=SUMIF(Data!A1:B3,"<>0",Data!B1)', '=AVERAGEIF(Data!A1:B3,"<>0",Data!B1)', '=COUNTIFS(Data!A1:A5,"=",Data!B1:B5,"<>")',
            '=SUMIF(Data!A:A,"",Data!B:B)', '=SUMIFS(Data!B:B,Data!A:A,"")', '=AVERAGEIF(Data!A:A,"",Data!B:B)', '=AVERAGEIFS(Data!B:B,Data!A:A,"")',
            '=IFERROR(SUMIF(Data!A1:A27,"#N/A",Data!B1),99)', '=COUNTIFS(Data!A1:A27,"<>0",Data!A1:A27,"<10")',
            '=SUMIF(Data!A:A,"",Data!B2)', '=AVERAGEIF(Data!A:A,"",Data!B2)', '=COUNTIF(Data!E1:E5,"2,5")', '=COUNTIF(Data!E1:E5,"1,000")',
            '=COUNTIF(Data!E1:E5,"2.5%")', '=COUNTIF(Data!E1:E5,"1000")', '=COUNTIF(Data!E6:E7,"a~~b")', '=COUNTIF(Data!E6:E7,"a~~*")',
            '=COUNTIF(Data!A1:A27,">#DIV/0!")', '=COUNTIF(Data!A1:A27,"<#N/A")', '=COUNTIF(Data!A1:A27,"<>")', '=COUNTIF(Data!A1:A27,"=*")',
            '=COUNTIF(Data!A1:A27,">")', '=COUNTIF(Data!A1:A27,"<")', '=COUNTIF(Data!A1:A27,">=")', '=COUNTIF(Data!A1:A27,"<=")',
            '=COUNTIF(Data!A1:A27,"1*")', '=COUNTIF(Data!A1:A27,"1E0")', '=COUNTIF(Data!A1:A27,"+1")', '=COUNTIF(Data!A1:A27,"= ")',
            '=SUMIF(Data!A1:A27,1/0,Data!B1)', '=COUNTIF(Data!A1:A27,"<>#VALUE!")', '=COUNTIF(Data!A1:A27,Data!D1:D1)')
        for ($i=0; $i -lt $extra.Count; $i++) { Add-Case "edge-$i" $extra[$i] }
        $book=$excel.Workbooks.Add(); while ($book.Worksheets.Count -gt 1) { $extraSheet=$book.Worksheets.Item($book.Worksheets.Count); $extraSheet.Delete(); Release-Com $extraSheet }
        $data=$book.Worksheets.Item(1)
        try { $data.Name='Data'; foreach ($item in $inputs) { Put $data $item }; $result=$book.Worksheets.Add([Type]::Missing,$data) } finally { Release-Com $data }
        try {
            $result.Name='Results'; foreach ($case in $cases) { Put $result @{Address=$case.Address;Kind='Formula';Value=$case.Formula} }
            foreach ($formula in @('=COUNTIF(1,1)', '=COUNTBLANK(0)', '=SUMIF(Data!A1:A2,1,5)', '=AVERAGEIF(Data!A1:A2,1,5)')) {
                try { Put $result @{Address='Z1';Kind='Formula';Value=$formula}; throw 'Excel accepted an expected-invalid range argument.' }
                catch { if ($_.Exception.Message -notlike 'Cannot set *') { throw }; $rejected += [ordered]@{Formula=$formula;Reason=$_.Exception.Message} }
            }
        } finally { Release-Com $result }
    }
    $excel.CalculateFullRebuild(); $initial=Capture $book $cases
    $saved=Join-Path $target 'initial.xlsx'; $book.SaveAs($saved,51); $book.Close($false); Release-Com $book; $book=$null
    $book=$excel.Workbooks.Open($saved,0,$true); $excel.CalculateFullRebuild(); $reopened=Capture $book $cases
    $stable=($initial|ConvertTo-Json -Depth 5 -Compress) -ceq ($reopened|ConvertTo-Json -Depth 5 -Compress)
    $edited=$null; $editedStable=$null
    if (-not $WorkbookPath) {
        $data=$book.Worksheets.Item('Data'); try { Put $data @{Address='A1';Kind='Number';Value=7}; Put $data @{Address='B3';Kind='Number';Value=300} } finally { Release-Com $data }
        $excel.CalculateFullRebuild(); $edited=Capture $book $cases
        $book.SaveAs((Join-Path $target 'edited.xlsx'),51); $book.Close($false); Release-Com $book; $book=$null
        $book=$excel.Workbooks.Open((Join-Path $target 'edited.xlsx'),0,$true); $excel.CalculateFullRebuild(); $editedReopened=Capture $book $cases
        $editedStable=($edited|ConvertTo-Json -Depth 5 -Compress) -ceq ($editedReopened|ConvertTo-Json -Depth 5 -Compress)
    }
    $book.Close($false); Release-Com $book; $book=$null
    $unchanged=if($WorkbookPath){$sourceHash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}else{$true}
    $matches=if($WorkbookPath){($oracle.$State|ConvertTo-Json -Depth 5 -Compress) -ceq ($initial|ConvertTo-Json -Depth 5 -Compress)}else{$null}
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=[int]$excel.Calculation;
        Inputs=$inputs;Cases=$cases;Rejected=$rejected;Initial=$initial;Edited=$edited;Stable=$stable;EditedStable=$editedStable;SourceUnchanged=$unchanged;SourceSha256=$sourceHash;ReferenceMatches=$matches;
        InitialSha256=(Get-FileHash -LiteralPath $saved -Algorithm SHA256).Hash} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $target 'conditional-functions.json') -Encoding utf8
    if (-not $stable -or $editedStable -eq $false -or -not $unchanged -or $matches -eq $false) { throw 'Conditional Excel oracle/verification failed; see report.' }
    Write-Output "PASS: $($cases.Count) conditional formulas; calculation and save/reopen stable."
} finally {
    if ($book) { $book.Close($false); Release-Com $book }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excel) { $excel.Quit(); Release-Com $excel }
}
