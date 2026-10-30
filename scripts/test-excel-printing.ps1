param([Parameter(Mandatory)][string]$OutputDirectory, [string]$WorkbookPath, [string]$ReferencePath,
    [ValidateSet('AfterReopen','AfterInsert','AfterDelete')][string]$ReferenceState='AfterReopen')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($object){if($null -ne $object){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}}
function Capture-Printing($book){
    $records=@()
    for($i=1;$i -le $book.Worksheets.Count;$i++){
        $sheet=$book.Worksheets.Item($i);$setup=$sheet.PageSetup
        try{
            $record=[ordered]@{Name=[string]$sheet.Name}
            foreach($name in @('PaperSize','Orientation','Zoom','FitToPagesWide','FitToPagesTall','Order','FirstPageNumber','BlackAndWhite','Draft','PrintComments','PrintErrors',
                'CenterHorizontally','CenterVertically','PrintHeadings','PrintGridlines','PrintArea','PrintTitleRows','PrintTitleColumns',
                'LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter',
                'DifferentFirstPageHeaderFooter','OddAndEvenPagesHeaderFooter','ScaleWithDocHeaderFooter','AlignMarginsHeaderFooter')){
                $record[$name]=$setup.$name
            }
            foreach($name in @('LeftMargin','RightMargin','TopMargin','BottomMargin','HeaderMargin','FooterMargin')){$record[$name]=[Math]::Round([double]$setup.$name,8)}
            foreach($pageName in @('FirstPage','EvenPage')){
                $page=$setup.$pageName
                try{foreach($name in @('LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter')){
                    $section=$page.$name;try{$record[$pageName+$name]=[string]$section.Text}finally{Release-Com $section}
                }}finally{Release-Com $page}
            }
            foreach($axis in @('H','V')){
                $breaks=$sheet.($axis+'PageBreaks');$positions=@()
                try{for($b=1;$b -le $breaks.Count;$b++){
                    $item=$breaks.Item($b)
                    try{if($item.Type -eq -4135){$location=$item.Location;try{$positions+=if($axis -eq 'H'){[int]$location.Row}else{[int]$location.Column}}finally{Release-Com $location}}}finally{Release-Com $item}
                }}finally{Release-Com $breaks}
                $record[$axis+'ManualBreaks']=$positions
            }
            $records += $record
        }finally{Release-Com $setup;Release-Com $sheet}
    }
    return ,$records
}
$excel=$null;$book=$null
try{
    $sourceHash=if($WorkbookPath){(Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$null}
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    if($WorkbookPath){$book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath,$repo),0,$true)}
    else{
        $book=$excel.Workbooks.Add()
        while($book.Worksheets.Count -gt 1){$extra=$book.Worksheets.Item($book.Worksheets.Count);try{$extra.Delete()}finally{Release-Com $extra}}
        for($i=0;$i -lt 3;$i++){
            if($i -eq 0){$sheet=$book.Worksheets.Item(1)}else{$last=$book.Worksheets.Item($book.Worksheets.Count);try{$sheet=$book.Worksheets.Add([Type]::Missing,$last)}finally{Release-Com $last}}
            $setup=$sheet.PageSetup
            try{
                $sheet.Name=@("报价,'版本",'适页报表','默认页面')[$i]
                $range=$sheet.Range('A1:J90');try{$range.Value2=[double]12}finally{Release-Com $range}
                if($i -eq 2){continue}
                $setup.PaperSize=9;$setup.Orientation=if($i -eq 0){2}else{1}
                $setup.LeftMargin=28.8;$setup.RightMargin=36;$setup.TopMargin=43.2;$setup.BottomMargin=50.4;$setup.HeaderMargin=14.4;$setup.FooterMargin=21.6
                $setup.Order=2;$setup.FirstPageNumber=3;$setup.BlackAndWhite=$true;$setup.PrintErrors=1
                $setup.CenterHorizontally=$true;$setup.CenterVertically=$false;$setup.PrintHeadings=$true;$setup.PrintGridlines=$true
                $setup.PrintArea=if($i -eq 0){'$A$1:$F$45,$H$1:$J$20'}else{'$A$1:$J$90'}
                $setup.PrintTitleRows='$1:$2';$setup.PrintTitleColumns='$A:$B'
                $setup.LeftHeader='&"Arial,Bold"中文报价';$setup.CenterHeader='&A';$setup.RightHeader='&P/&N'
                $setup.LeftFooter='&F';$setup.CenterFooter='&D &T';$setup.RightFooter='审核'
                $setup.DifferentFirstPageHeaderFooter=$true;$setup.OddAndEvenPagesHeaderFooter=$true;$setup.ScaleWithDocHeaderFooter=$false;$setup.AlignMarginsHeaderFooter=$true
                foreach($pageName in @('FirstPage','EvenPage')){$page=$setup.$pageName;try{$section=$page.CenterHeader;try{$section.Text=$pageName+' 标题'}finally{Release-Com $section};$section=$page.RightFooter;try{$section.Text=$pageName+' 尾注'}finally{Release-Com $section}}finally{Release-Com $page}}
                if($i -eq 0){$setup.Zoom=75}else{$setup.Zoom=$false;$setup.FitToPagesWide=1;$setup.FitToPagesTall=2}
                $range=$sheet.Range('A21');try{$breaks=$sheet.HPageBreaks;try{$item=$breaks.Add($range);Release-Com $item}finally{Release-Com $breaks}}finally{Release-Com $range}
                $range=$sheet.Range('D1');try{$breaks=$sheet.VPageBreaks;try{$item=$breaks.Add($range);Release-Com $item}finally{Release-Com $breaks}}finally{Release-Com $range}
            }finally{Release-Com $setup;Release-Com $sheet}
        }
    }
    $before=Capture-Printing $book;$calculation=[int]$excel.Calculation
    $saved=Join-Path $target 'excel-source.xlsx';$book.SaveAs($saved,51)
    $book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($saved,0,$true);$after=Capture-Printing $book
    $insert=$null;$deleted=$null
    if(-not $WorkbookPath){
        $sheet=$book.Worksheets.Item(1)
        try{
            $sheet.Name="改名'打印";$range=$sheet.Rows.Item(1);try{[void]$range.Insert()}finally{Release-Com $range}
            $range=$sheet.Columns.Item(1);try{[void]$range.Insert()}finally{Release-Com $range}
        }finally{Release-Com $sheet}
        $insert=Capture-Printing $book;$book.SaveAs((Join-Path $target 'excel-inserted.xlsx'),51)
        $book.Close($false);Release-Com $book;$book=$excel.Workbooks.Open($saved,0,$true)
        $sheet=$book.Worksheets.Item(1)
        try{
            $range=$sheet.Range('1:2');try{[void]$range.Delete()}finally{Release-Com $range}
            $range=$sheet.Range('A:B');try{[void]$range.Delete()}finally{Release-Com $range}
        }finally{Release-Com $sheet}
        $deleted=Capture-Printing $book;$book.SaveAs((Join-Path $target 'excel-deleted.xlsx'),51)
    }
    $book.Close($false);Release-Com $book;$book=$null
    $unchanged=if($WorkbookPath){$sourceHash -eq (Get-FileHash -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath,$repo))).Hash}else{$true}
    $same=($before|ConvertTo-Json -Depth 8 -Compress) -ceq ($after|ConvertTo-Json -Depth 8 -Compress)
    $matches=if($ReferencePath){$reference=Get-Content -LiteralPath $ReferencePath -Raw|ConvertFrom-Json; ($reference.$ReferenceState|ConvertTo-Json -Depth 8 -Compress) -ceq ($before|ConvertTo-Json -Depth 8 -Compress)}else{$null}
    $report=[ordered]@{ExcelVersion=[string]$excel.Version;ExcelBuild=[string]$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);Calculation=$calculation;
        Before=$before;AfterReopen=$after;AfterInsert=$insert;AfterDelete=$deleted;SameAfterReopen=$same;ReferenceMatches=$matches;SourceUnchanged=$unchanged;SourceHash=$sourceHash}
    $report|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $target 'excel-printing.json')
    if(-not $same -or -not $unchanged -or $matches -eq $false){throw 'Excel printing comparison failed; see JSON.'}
    Write-Output "PASS: Excel printing settings and save/reopen; $target"
}finally{if($book){$book.Close($false);Release-Com $book};if($excel){[GC]::Collect();[GC]::WaitForPendingFinalizers();$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
