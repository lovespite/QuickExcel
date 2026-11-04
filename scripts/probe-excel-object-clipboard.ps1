[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$run=Join-Path $root ('artifacts/clipboard/objects-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $run)
Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardNative.cs')
Add-Type -AssemblyName System.Drawing
function Release-Com($value) {if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
function Capture([string]$name) {
    $dir=Join-Path $run $name;[void](New-Item -ItemType Directory -Path $dir)
    $data=[ClipboardNative]::Read();$formats=@()
    foreach($key in $data.Keys) {
        $formats+=@{Name=$key;Bytes=$data[$key].Length}
        if($data[$key].Length){[IO.File]::WriteAllBytes((Join-Path $dir (($key -replace '[^a-zA-Z0-9]','_')+'.bin')),$data[$key])}
    }
    $formats|ConvertTo-Json|Set-Content (Join-Path $dir 'formats.json') -Encoding utf8
    Write-Output $name;Write-Output ($formats|ConvertTo-Json -Compress)
}
$excel=$null;$book=$null;$sheet=$null;$chartObject=$null;$chart=$null;$picture=$null
try {
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
    $matrix=[object[,]]::new(4,3)
    $matrix[0,0]='Category';$matrix[0,1]='Revenue';$matrix[0,2]='Cost'
    for($r=1;$r -lt 4;$r++){$matrix[$r,0]='Q'+$r;$matrix[$r,1]=[double](10*$r);$matrix[$r,2]=[double](7*$r)}
    $range=$sheet.Range('A1:C4');$range.Value2=$matrix
    $objects=$sheet.ChartObjects();$chartObject=$objects.Add(100,100,480,300);Release-Com $objects
    $chart=$chartObject.Chart;$chart.ChartType=51;$chart.SetSourceData($range);$chart.HasTitle=$true;$chart.ChartTitle.Text='Clipboard Chart';Release-Com $range
    $book.SaveAs((Join-Path $run 'chart.xlsx'),51)
    $chartObject.Copy();Capture 'chart'
    $chart.CopyPicture(1,2);Capture 'chart-picture'
    $bitmap=[Drawing.Bitmap]::new(80,40);$graphics=[Drawing.Graphics]::FromImage($bitmap)
    try{$graphics.Clear([Drawing.Color]::CornflowerBlue);$bitmap.Save((Join-Path $run 'fixture.png'),[Drawing.Imaging.ImageFormat]::Png)}finally{$graphics.Dispose();$bitmap.Dispose()}
    $picture=$sheet.Shapes.AddPicture((Join-Path $run 'fixture.png'),0,-1,10,10,80,40)
    $picture.Copy();Capture 'image'
    Write-Output "Evidence: $run"
} finally {
    Release-Com $picture;Release-Com $chart;Release-Com $chartObject;Release-Com $sheet
    if($book){$book.Close($false);Release-Com $book};if($excel){$excel.Quit();Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
