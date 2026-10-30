param([Parameter(Mandatory)][string]$OutputDirectory)
# Reproduce the embedded table-style projection from an owned Excel workbook.
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$excel=$null;$book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
 if(-not ('TableCatalogOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class TableCatalogOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
 [uint32]$ownedExcelProcess=0;[void][TableCatalogOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$ownedExcelProcess)
 $evidence=@{Version=$excel.Version;Build=$excel.Build;Language=$excel.LanguageSettings.LanguageID(2);ProcessId=$ownedExcelProcess;Started=(Get-Process -Id $ownedExcelProcess).StartTime}
 $book=$excel.Workbooks.Add();$styles=$book.TableStyles
 try{foreach($family in @(@('Light',21),@('Medium',28),@('Dark',11))){for($i=1;$i -le $family[1];$i++){
  $name='TableStyle'+$family[0]+$i;$style=$styles.Item($name)
  try{$copy=$style.Duplicate('QE'+$name);Release-Com $copy}finally{Release-Com $style}
 }}}finally{Release-Com $styles}
 $path=Join-Path $target 'catalog.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
 $evidence.Sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
 $evidence|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
}finally{if($book){$book.Close($false);Release-Com $book};if($excel){[GC]::Collect();[GC]::WaitForPendingFinalizers();$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=[IO.Compression.ZipFile]::OpenRead($path)
try{
 $reader=[IO.StreamReader]::new($zip.GetEntry('xl/styles.xml').Open())
 try{[xml]$source=$reader.ReadToEnd()}finally{$reader.Dispose()}
 $ns=[Xml.XmlNamespaceManager]::new($source.NameTable);$uri='http://schemas.openxmlformats.org/spreadsheetml/2006/main';$ns.AddNamespace('s',$uri)
 $dxfs=$source.SelectNodes('/s:styleSheet/s:dxfs/s:dxf',$ns)
 $catalog=[Xml.XmlDocument]::new();$root=$catalog.CreateElement('tableStyleCatalog');$root.SetAttribute('xmlns:s',$uri);[void]$catalog.AppendChild($root)
 foreach($style in ($source.SelectNodes('/s:styleSheet/s:tableStyles/s:tableStyle',$ns)|Sort-Object { $_.GetAttribute('name') })){
  $name=$style.GetAttribute('name');if(-not $name.StartsWith('QETableStyle')){continue}
  $item=$catalog.CreateElement('style');$item.SetAttribute('name',$name.Substring(2));[void]$root.AppendChild($item)
  foreach($part in $style.SelectNodes('s:tableStyleElement',$ns)){
   $element=$catalog.CreateElement('element');$element.SetAttribute('type',$part.GetAttribute('type'))
   $size=$part.GetAttribute('size');if(-not $size){$size='1'};$element.SetAttribute('size',$size)
   $dxf=$catalog.ImportNode($dxfs[[int]$part.GetAttribute('dxfId')],$true)
   foreach($node in @($dxf)+@($dxf.SelectNodes('.//*'))){$node.Prefix='s'}
   [void]$element.AppendChild($dxf);[void]$item.AppendChild($element)
  }
 }
 if($root.ChildNodes.Count -ne 60){throw 'Expected all 60 built-in table styles.'}
 $settings=[Xml.XmlWriterSettings]::new();$settings.Indent=$true;$settings.Encoding=[Text.UTF8Encoding]::new($false)
 $writer=[Xml.XmlWriter]::Create((Join-Path $target 'BuiltinTableStyles.xml'),$settings)
 try{$catalog.Save($writer)}finally{$writer.Dispose()}
 $stream=$zip.GetEntry('xl/theme/theme1.xml').Open();$output=[IO.File]::Create((Join-Path $target 'DefaultTableTheme.xml'))
 try{$stream.CopyTo($output)}finally{$output.Dispose();$stream.Dispose()}
}finally{$zip.Dispose()}
