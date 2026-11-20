[CmdletBinding()]
param([Parameter(Mandatory)][string]$CertificateThumbprint,
      [string]$OutputDirectory = '', [switch]$UseActivePackage)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot/windows-signing.ps1"
if (-not $IsWindows) { throw 'ClickOnce publishing requires Windows.' }
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts = Join-Path $root 'artifacts'
function ArtifactPath([string]$value) {
    $path = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($value)) { $value } else { Join-Path $root $value }))
    if (-not $path.StartsWith($artifacts + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Output must stay under this repository artifacts directory.' }
    for ($parent = $path; $parent; $parent = Split-Path $parent -Parent) {
        if ((Test-Path -LiteralPath $parent) -and ((Get-Item -LiteralPath $parent).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Output cannot traverse reparse points.' }
    }
    $path
}
function Invoke-Mage([string[]]$Arguments) {
    & $mage @Arguments
    if ($LASTEXITCODE -ne 0) { throw "dotnet-mage failed: $($Arguments[0])" }
}
$certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$CertificateThumbprint"
if (-not $certificate.HasPrivateKey) { throw 'The signing private key is missing.' }
$version = ([xml](Get-Content -LiteralPath (Join-Path $root 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj') -Raw)).SelectSingleNode('/Project/PropertyGroup/Version').InnerText
if ($version -notmatch '^\d+\.\d+\.\d+(\.\d+)?$') { throw 'ClickOnce requires a numeric release version.' }
$numericVersion = [version]$version
$deploymentVersion = '{0}.{1}.{2}.{3}' -f $numericVersion.Major,$numericVersion.Minor,$numericVersion.Build,([Math]::Max(0,$numericVersion.Revision))
$runId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8)
$output = ArtifactPath $(if ($OutputDirectory) { $OutputDirectory } else { "artifacts/clickonce/QuickExcel-$version-win-x64-$runId" })
if (Test-Path -LiteralPath $output) { throw 'Publish into a new directory; existing packages are preserved.' }
$work = ArtifactPath "artifacts/clickonce/work/$runId"
$toolPath = ArtifactPath 'artifacts/tools/dotnet-mage'
$mage = Join-Path $toolPath 'dotnet-mage.exe'
if (-not (Test-Path -LiteralPath $mage)) {
    & dotnet tool install Microsoft.DotNet.Mage --version 10.0.0 --tool-path $toolPath
    if ($LASTEXITCODE -ne 0) { throw 'Installing pinned dotnet-mage failed.' }
}
$magePackage = Join-Path $toolPath '.store/microsoft.dotnet.mage/10.0.0/microsoft.dotnet.mage/10.0.0'
if (-not (Test-Path -LiteralPath (Join-Path $magePackage 'Microsoft.DotNet.Mage.nuspec'))) { throw 'Expected dotnet-mage 10.0.0.' }
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$msbuild = @(& $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe') | Select-Object -First 1
if (-not $msbuild) { throw 'Visual Studio MSBuild is required for the ClickOnce setup bootstrapper.' }
$bootstrapperSdk = Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs/ClickOnce Bootstrapper'
if (-not (Test-Path -LiteralPath (Join-Path $bootstrapperSdk 'Engine/setup.bin'))) { throw 'Install the Visual Studio ClickOnce bootstrapper SDK.' }
$download = ArtifactPath 'artifacts/clickonce/downloads/MicrosoftEdgeWebview2Setup.exe'
$downloadUrl = 'https://go.microsoft.com/fwlink/p/?LinkId=2124703'
New-Item -ItemType Directory -Path (Split-Path $download -Parent),$work -Force | Out-Null
if (-not (Test-Path -LiteralPath $download)) {
    Invoke-WebRequest $downloadUrl -OutFile ($download + '.partial')
    Assert-QuickExcelWebView2Bootstrapper ($download + '.partial')
    Move-Item -LiteralPath ($download + '.partial') -Destination $download
}
Assert-QuickExcelWebView2Bootstrapper $download
$webviewSignature = Get-AuthenticodeSignature -LiteralPath $download
if (-not $UseActivePackage) {
    & "$PSScriptRoot/build-quickexcel.ps1" -SkipTests -RunSmoke -CodeSigningCertificateThumbprint $CertificateThumbprint
}
$pointer = Get-Content -LiteralPath (Join-Path $artifacts 'avalonia/active-windows.json') -Raw | ConvertFrom-Json
$source = ArtifactPath $pointer.directory
$releasePath = Join-Path $source 'release.json'
if ((Get-FileHash -LiteralPath $releasePath -Algorithm SHA256).Hash -ne $pointer.manifestSha256) { throw 'Active package manifest hash mismatch.' }
$release = Get-Content -LiteralPath $releasePath -Raw | ConvertFrom-Json
if ($release.version -ne $version -or -not $release.checks.passed -or -not $release.selfContained -or $release.rid -ne 'win-x64') { throw 'Active package is not the verified self-contained Windows build of the current version.' }
foreach ($file in $release.files) {
    $path = Join-Path $source $file.path
    if ((Get-Item -LiteralPath $path).Length -ne $file.size -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $file.sha256) { throw "Active payload mismatch: $($file.path)" }
}
$folder = 'Application Files/QuickExcel_' + $deploymentVersion.Replace('.','_')
$payload = Join-Path $output $folder
New-Item -ItemType Directory -Path $payload -Force | Out-Null
Get-ChildItem -LiteralPath $source -Force | Copy-Item -Destination $payload -Recurse
foreach ($file in @(Get-ChildItem -LiteralPath $payload -Recurse -File | Where-Object { $_.Name -match '^QuickExcel\..*\.(exe|dll)$|^quickxl\.exe$' })) {
    Assert-QuickExcelCodeSignature $file.FullName $CertificateThumbprint
}
Copy-Item -LiteralPath (Join-Path $root 'QuickExcel.Desktop.Avalonia/Assets/QuickExcel.ico') -Destination (Join-Path $payload 'QuickExcel.ico')
Invoke-Mage -Arguments @('-AddLauncher','quickxl.exe','-TargetDirectory',$payload)
& "$PSScriptRoot/sign-windows-code.ps1" -CertificateThumbprint $CertificateThumbprint -Path (Join-Path $payload 'Launcher.exe')
$appManifest = Join-Path $payload 'QuickExcel.manifest'
Invoke-Mage -Arguments @('-New','Application','-ToFile',$appManifest,'-Name','QuickExcel','-Version',$deploymentVersion,'-Processor','msil','-FromDirectory',$payload,'-IconFile','QuickExcel.ico','-Algorithm','sha256RSA')
[xml]$applicationXml = Get-Content -LiteralPath $appManifest -Raw
$applicationXml.SelectSingleNode('//*[local-name()="dependentOS"]//*[local-name()="os"]').SetAttribute('majorVersion','10')
$applicationXml.SelectSingleNode('//*[local-name()="dependentOS"]//*[local-name()="os"]').SetAttribute('minorVersion','0')
$applicationXml.Save($appManifest)
Invoke-Mage -Arguments @('-Sign',$appManifest,'-CertHash',$CertificateThumbprint,'-Algorithm','sha256RSA')
$deployment = Join-Path $output 'QuickExcel.application'
Invoke-Mage -Arguments @('-New','Deployment','-ToFile',$deployment,'-Name','QuickExcel','-Publisher','QuickExcel','-Version',$deploymentVersion,'-Processor','msil','-Install','true','-IncludeProviderURL','false','-AppManifest',$appManifest,'-AppCodeBase',($folder + '/QuickExcel.manifest'),'-Algorithm','sha256RSA')
# Manual application updates do not use Mage's before-startup update check.
[xml]$manifest = Get-Content -LiteralPath $deployment -Raw
$manifest.SelectSingleNode('//*[local-name()="compatibleFrameworks"]/*[local-name()="framework"]').SetAttribute('targetVersion','4.8')
foreach ($node in @($manifest.SelectNodes('//*[local-name()="subscription" or local-name()="deploymentProvider"]'))) { [void]$node.ParentNode.RemoveChild($node) }
$manifest.Save($deployment)
Invoke-Mage -Arguments @('-Sign',$deployment,'-CertHash',$CertificateThumbprint,'-Algorithm','sha256RSA')
$bootstrapper = Join-Path $work 'bootstrapper'
New-Item -ItemType Directory -Path (Join-Path $bootstrapper 'Packages') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $bootstrapperSdk 'Engine') -Destination $bootstrapper -Recurse
Copy-Item -LiteralPath (Join-Path $bootstrapperSdk 'Schemas') -Destination $bootstrapper -Recurse
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'clickonce/WebView2') -Destination (Join-Path $bootstrapper 'Packages') -Recurse
Copy-Item -LiteralPath $download -Destination (Join-Path $bootstrapper 'Packages/WebView2')
& $msbuild (Join-Path $PSScriptRoot 'clickonce/bootstrapper.proj') /nologo /v:minimal "/p:BootstrapperRoot=$bootstrapper" "/p:InstallerOutput=$output"
if ($LASTEXITCODE -ne 0) { throw 'Creating setup.exe failed.' }
& "$PSScriptRoot/sign-windows-code.ps1" -CertificateThumbprint $CertificateThumbprint -Path (Join-Path $output 'setup.exe')
Export-Certificate -Cert $certificate -FilePath (Join-Path $output 'QuickExcel-CodeSigning.cer') | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'clickonce/README.txt') -Destination (Join-Path $output 'README.txt')
& "$PSScriptRoot/test-clickonce.ps1" -PackageDirectory $output -CertificateThumbprint $CertificateThumbprint -EvidenceDirectory (Join-Path $work 'verification')
$files = @(Get-ChildItem -LiteralPath $output -File -Recurse | Sort-Object FullName | ForEach-Object { [ordered]@{ path=[IO.Path]::GetRelativePath($output,$_.FullName).Replace('\','/'); size=$_.Length; sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
[ordered]@{ schemaVersion=1; productVersion=$version; deploymentVersion=$deploymentVersion; rid='win-x64'; offline=$false; automaticUpdates=$false; certificateThumbprint=$CertificateThumbprint; certificateNotAfter=$certificate.NotAfter.ToString('o'); sourcePackage=$source; sourceManifestSha256=$pointer.manifestSha256; mageVersion='10.0.0'; webView2=[ordered]@{ distribution='EvergreenBootstrapper'; internetRequiredWhenMissing=$true; url=$downloadUrl; size=(Get-Item -LiteralPath $download).Length; version=(Get-Item -LiteralPath $download).VersionInfo.ProductVersion; sha256=(Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash; signer=$webviewSignature.SignerCertificate.Subject }; evidence=(Join-Path $work 'verification'); files=$files } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $output 'clickonce-release.json') -Encoding utf8
$archive = $output + '.zip'
[IO.Compression.ZipFile]::CreateFromDirectory($output,$archive,[IO.Compression.CompressionLevel]::Optimal,$false)
(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash + '  ' + [IO.Path]::GetFileName($archive) | Set-Content -LiteralPath ($archive+'.sha256') -Encoding ascii
Write-Host "PASS: ClickOnce $deploymentVersion; $output; $archive"
