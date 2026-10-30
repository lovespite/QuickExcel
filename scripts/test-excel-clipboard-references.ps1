param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$expected=@('G5:H6','G$5:H6','$G5:H6','$G$5:H6','G5:H$6','$A3:$B4','A$3:$B4','$A$3:$B4',
    'G5:$H6','$A3:B$4','A$3:B$4','$A$3:B$4','G5:$H$6','$A3:$B$4','A$3:$B$4','$A$3:$B$4')
$excel=$null;$book=$null;$results=@()
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
    $source=$sheet.Range('A1:C2');$destination=$sheet.Range('E5');$cell=$sheet.Range('C1');$output=$sheet.Range('E7')
    try{for($flags=0;$flags -lt 16;$flags++){
        $reference=('{0}A{1}3:{2}B{3}4' -f $(if($flags -band 1){'$'}else{''}),$(if($flags -band 2){'$'}else{''}),
            $(if($flags -band 4){'$'}else{''}),$(if($flags -band 8){'$'}else{''}))
        $cell.Formula='=SUM('+$reference+')';[void]$source.Copy();[void]$destination.PasteSpecial(-4104,-4142,$false,$true)
        $wanted='=SUM('+$expected[$flags]+')';$actual=[string]$output.Formula
        $results+=@{Flags=$flags;Source=$cell.Formula;Expected=$wanted;Actual=$actual;Passed=($wanted -ceq $actual)}
    }
    $book.SaveAs((Join-Path $target 'transpose-reference.xlsx'),51)
    }finally{Release-Com $output;Release-Com $cell;Release-Com $destination;Release-Com $source;Release-Com $sheet}
    @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Source='A1:C2';Destination='E5';Cases=$results}|
        ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $target 'transpose-reference.json') -Encoding utf8
    if(@($results|Where-Object {-not $_.Passed}).Count){throw 'Excel transpose reference behavior changed.'}
    Write-Output '16 independent Excel transpose reference cases passed.'
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
}
