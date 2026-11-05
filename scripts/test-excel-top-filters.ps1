param([Parameter(Mandatory)][string]$OutputDirectory, [string]$InputDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$output = [IO.Path]::GetFullPath($OutputDirectory, $repo)
$prefix = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts')) + [IO.Path]::DirectorySeparatorChar
if (-not $output.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $output)) { throw 'Use a new artifacts directory.' }
[void](New-Item -ItemType Directory -Path $output)
function Release-Com($value) { if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
function Visible-Rows($sheet) { @(2..10 | Where-Object { $row=$sheet.Rows.Item($_); try {-not $row.Hidden} finally {Release-Com $row} }) }
$specs=@(
    @{name='top2';count=2;percent=$false;bottom=$false;group=$false;table=$false},
    @{name='top3';count=3;percent=$false;bottom=$false;group=$false;table=$false},
    @{name='bottom2';count=2;percent=$false;bottom=$true;group=$false;table=$false},
    @{name='top50';count=50;percent=$true;bottom=$false;group=$false;table=$false},
    @{name='bottom50';count=50;percent=$true;bottom=$true;group=$false;table=$false},
    @{name='top1percent';count=1;percent=$true;bottom=$false;group=$false;table=$false},
    @{name='group-top2';count=2;percent=$false;bottom=$false;group=$true;table=$false},
    @{name='table-top2';count=2;percent=$false;bottom=$false;group=$false;table=$true}
)
$excel=$null; $book=$null; $sheet=$null; $cases=@()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false; $excel.AutomationSecurity=3
    foreach($spec in $specs) {
        $book=$excel.Workbooks.Add(-4167); $sheet=$book.Worksheets.Item(1)
        $sheet.Range('A1').Value2='Amount'; $sheet.Range('B1').Value2='Group'
        $values=@(10,20,20,30,40)
        for($i=0;$i -lt 5;$i++) { $row=$i+2; $cell=$sheet.Cells.Item($row,1); try {$cell.Value2=[double]$values[$i]} finally {Release-Com $cell} }
        $sheet.Range('A7').NumberFormat='@'; $sheet.Range('A7').Value2='999'
        $sheet.Range('A9').Value2='tail'; $sheet.Range('A10').Value2=$true
        $sheet.Range('B2:B4').Value2='A'; $sheet.Range('B5:B10').Value2='B'
        if($spec.table) { $list=$sheet.ListObjects.Add(1,$sheet.Range('A1:B10'),[Type]::Missing,1); Release-Com $list }
        if($spec.group) { [void]$sheet.Range('A1:B10').AutoFilter(2,'A') }
        $operator=if($spec.percent){if($spec.bottom){6}else{5}}else{if($spec.bottom){4}else{3}}
        $criterion=[string]$spec['count']; $filterOperator=[int]$operator
        try { [void]$sheet.Range('A1:B10').AutoFilter(1,$criterion,$filterOperator) }
        catch { throw "Excel rejected $($spec.name), count=$criterion, operator=$filterOperator : $($_.Exception.Message)" }
        $visible=@(Visible-Rows $sheet)
        $path=Join-Path $output ('excel-'+$spec.name+'.xlsx'); $book.SaveAs($path,51)
        $book.Close($false); Release-Com $sheet; $sheet=$null; Release-Com $book; $book=$null
        $zip=[IO.Compression.ZipFile]::OpenRead($path)
        try {
            $parts=[ordered]@{}
            foreach($entry in $zip.Entries) { if($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')) { $reader=[IO.StreamReader]::new($entry.Open());try {$parts[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()} } }
        }finally{$zip.Dispose()}
        if($InputDirectory) {
            $source=Join-Path ([IO.Path]::GetFullPath($InputDirectory,$repo)) ($spec.name+'.xlsx')
            $book=$excel.Workbooks.Open($source,0,$true);$sheet=$book.Worksheets.Item(1)
            if($spec.table){$sheet.ListObjects.Item(1).AutoFilter.ApplyFilter()}else{$sheet.AutoFilter.ApplyFilter()}
            $actual=@(Visible-Rows $sheet)
            if(($actual -join ',') -ne ($visible -join ',')){throw ($spec.name+' reapply differs: '+($actual -join ','))}
            $book.SaveAs((Join-Path $output ('roundtrip-'+$spec.name+'.xlsx')),51)
            $book.Close($false); Release-Com $sheet; $sheet=$null; Release-Com $book; $book=$null
        }
        $cases += [ordered]@{spec=$spec;visibleRows=$visible;sourceSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;parts=$parts}
    }
    $rejections=[ordered]@{}
    foreach($kind in @('error','noNumbers')) {
        $book=$excel.Workbooks.Add(-4167);$sheet=$book.Worksheets.Item(1);$sheet.Range('A1').Value2='Header'
        if($kind -eq 'error'){$sheet.Range('A2').Value2=10;$sheet.Range('A3').Formula='=1/0'}else{$sheet.Range('A2').NumberFormat='@';$sheet.Range('A2').Value2='999'}
        $rejected=$false
        try{[void]$sheet.Range('A1:A3').AutoFilter(1,2,3)}catch{$rejected=$true}
        if(-not $rejected){throw "Unexpectedly accepted $kind"};$rejections[$kind]=$rejected
        $book.Close($false);Release-Com $sheet;$sheet=$null;Release-Com $book;$book=$null
    }
    [ordered]@{excelVersion=$excel.Version;excelBuild=$excel.Build;cases=$cases;rejections=$rejections} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $output 'excel-top-filters.json') -Encoding utf8
    Write-Host "PASS: independent top-filter cases captured: $output"
}finally {if($null -ne $book){$book.Close($false)};if($null -ne $excel){$excel.Quit()};Release-Com $sheet;Release-Com $book;Release-Com $excel}
