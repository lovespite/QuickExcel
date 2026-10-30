param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$excel=$null;$book=$null;$cases=@()
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    foreach($case in @(@{Selection='B2';Command='FillDown'},@{Selection='A2:B2';Command='FillDown'},@{Selection='B2';Command='FillRight'},@{Selection='B1:B2';Command='FillRight'})){
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        try{
            $cell=$sheet.Range('A1');try{$cell.Value2=[double]3}finally{Release-Com $cell}
            $cell=$sheet.Range('A2');try{$cell.Value2=[double]7}finally{Release-Com $cell}
            $cell=$sheet.Range('B1');try{$cell.Formula='=A1+1'}finally{Release-Com $cell}
            $range=$sheet.Range($case.Selection);$bars=$excel.CommandBars
            try{[void]$range.Select();$bars.ExecuteMso($case.Command)}finally{Release-Com $bars;Release-Com $range}
            $excel.CalculateFullRebuild();$cells=@()
            foreach($address in @('A1','B1','A2','B2')){
                $cell=$sheet.Range($address)
                try{$cells+=@{Address=$address;Value=$cell.Value2;Formula=$cell.Formula}}finally{Release-Com $cell}
            }
            $cases+=@{Selection=$case.Selection;Command=$case.Command;Cells=$cells}
        }finally{Release-Com $sheet;$book.Close($false);Release-Com $book;$book=$null}
    }
    @{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Method='Owned hidden Excel CommandBars.ExecuteMso, not physical keyboard';Cases=$cases}|
        ConvertTo-Json -Depth 8|Set-Content (Join-Path $target 'excel-fill-selection.json') -Encoding utf8
}finally{
    if($book){$book.Close($false);Release-Com $book}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($excel){$excel.Quit();Release-Com $excel}
}
