[CmdletBinding()]
param([Parameter(Mandatory)][string]$PackageDirectory,
      [Parameter(Mandatory)][string]$CertificateThumbprint,
      [Parameter(Mandatory)][string]$EvidenceDirectory,
      [string]$ExpectedJscSourceUrl = '')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot/windows-signing.ps1"
$package = (Get-Item -LiteralPath $PackageDirectory).FullName
$evidence = [IO.Path]::GetFullPath($EvidenceDirectory)
$artifacts = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../artifacts'))
if (-not $evidence.StartsWith($artifacts + '\', [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $evidence)) { throw 'Use a new evidence directory under artifacts.' }
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
$certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$CertificateThumbprint"
Add-Type -AssemblyName System.Security.Cryptography.Xml
[Security.Cryptography.CryptoConfig]::AddAlgorithm([QuickExcel.Publishing.ClickOnceSha256SignatureDescription], 'http://www.w3.org/2000/09/xmldsig#rsa-sha256')
[Security.Cryptography.CryptoConfig]::AddAlgorithm([Security.Cryptography.SHA256Managed], 'http://www.w3.org/2000/09/xmldsig#sha256')
function Read-SignedManifest([string]$path) {
    $document = [xml]::new()
    $document.PreserveWhitespace = $true
    $document.Load($path)
    $signature = $document.DocumentElement.SelectSingleNode('*[local-name()="Signature" and namespace-uri()="http://www.w3.org/2000/09/xmldsig#"]')
    if (-not $signature) { throw "Missing manifest signature: $path" }
    $signedXml = [Security.Cryptography.Xml.SignedXml]::new($document)
    $signedXml.LoadXml($signature)
    if (-not $signedXml.CheckSignature($certificate, $true)) { throw "Manifest signature mismatch: $path" }
    return ,$document
}
function Check-Hash([System.Xml.XmlElement]$reference, [string]$baseDirectory, [string]$relative) {
    if ([IO.Path]::IsPathRooted($relative) -or $relative -match '(^|[/\\])\.\.([/\\]|$)') { throw 'Manifest paths must stay inside the package.' }
    $file = Join-Path $baseDirectory $relative
    $hash = $reference.SelectSingleNode('*[local-name()="hash"]')
    if (-not $hash -or $hash.SelectSingleNode('*[local-name()="DigestMethod"]').GetAttribute('Algorithm') -notin @('http://www.w3.org/2001/04/xmlenc#sha256','http://www.w3.org/2000/09/xmldsig#sha256')) { throw "SHA-256 required: $relative" }
    $expected = $hash.SelectSingleNode('*[local-name()="DigestValue"]').InnerText
    $actual = [Convert]::ToBase64String([Convert]::FromHexString((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash))
    if ($actual -ne $expected -or (Get-Item -LiteralPath $file).Length -ne [long]$reference.GetAttribute('size')) { throw "Manifest file hash/size mismatch: $relative" }
    $file
}
$deploymentPath = Join-Path $package 'QuickExcel.application'
$deployment = Read-SignedManifest $deploymentPath
if ($deployment.SelectNodes('//*[local-name()="subscription" or local-name()="deploymentProvider"]').Count -ne 0) { throw 'The manual-update installer must not poll an application update URL.' }
if ($deployment.SelectSingleNode('//*[local-name()="deployment"]').GetAttribute('install') -ne 'true') { throw 'Local application installation must be enabled.' }
$appReference = $deployment.SelectSingleNode('//*[local-name()="dependentAssembly" and @dependencyType="install"]')
$appPath = Check-Hash $appReference $package $appReference.GetAttribute('codebase')
$app = Read-SignedManifest $appPath
$payload = Split-Path $appPath -Parent
$identity = $app.DocumentElement.SelectSingleNode('*[local-name()="assemblyIdentity"]')
$version = $identity.GetAttribute('version')
if ($identity.GetAttribute('processorArchitecture') -ne 'msil') { throw 'The ClickOnce manifest must match the AnyCPU launcher.' }
if ($deployment.DocumentElement.SelectSingleNode('*[local-name()="assemblyIdentity"]').GetAttribute('version') -ne $version) { throw 'Deployment and application versions disagree.' }
$entry = $app.SelectSingleNode('//*[local-name()="entryPoint"]/*[local-name()="commandLine"]')
if (-not $entry -or $entry.GetAttribute('file') -ne 'Launcher.exe') { throw 'Modern .NET ClickOnce must use Launcher.exe.' }
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($reference in $app.SelectNodes('/*[local-name()="assembly"]/*[local-name()="file"] | /*[local-name()="assembly"]/*[local-name()="dependency"]/*[local-name()="dependentAssembly" and @dependencyType="install"]')) {
    $relative = if ($reference.LocalName -eq 'file') { $reference.GetAttribute('name') } else { $reference.GetAttribute('codebase') }
    $file = Check-Hash $reference $payload $relative
    if (-not $seen.Add([IO.Path]::GetFullPath($file))) { throw "Duplicate manifest file: $relative" }
}
foreach ($file in Get-ChildItem -LiteralPath $payload -File -Recurse) {
    if ($file.FullName -ne $appPath -and -not $seen.Contains($file.FullName)) { throw "Unlisted payload file: $($file.FullName)" }
}
foreach ($file in Get-ChildItem -LiteralPath $package -File -Recurse) {
    if ($file.Name -eq 'JSC-CORRESPONDING-SOURCE.zip') { throw 'JSC source archives must not enter the ClickOnce package.' }
    if ($file.Name -match '\.(pfx|pdb|map)$|password|WorkerFixture|RecoveryFixture|HostFixture|SandboxProbe|PlatformSmoke|TestProbe') { throw "Forbidden installer file: $($file.Name)" }
    if ($file.Name -match '^QuickExcel\..*\.(exe|dll)$|^(quickxl|Launcher|setup)\.exe$') {
        Assert-QuickExcelCodeSignature $file.FullName $CertificateThumbprint
    }
}
$sourceUrl = (Get-Content -LiteralPath (Join-Path $payload 'JSC-Licenses/JSC-SOURCE.txt') -Raw).Trim()
$sourceUri = $null
if (-not [Uri]::TryCreate($sourceUrl, [UriKind]::Absolute, [ref]$sourceUri) -or
    $sourceUri.Scheme -ne 'https' -or $sourceUri.UserInfo -or $sourceUrl -match '\s') { throw 'JSC-SOURCE.txt must contain one public HTTPS source address.' }
if ($ExpectedJscSourceUrl -and $sourceUrl -cne $ExpectedJscSourceUrl) { throw 'JSC source address differs from the requested release source.' }
foreach ($name in @('JSC-REBUILD.txt','JSC-THIRD-PARTY-NOTICES.txt')) {
    $notice = Get-Content -LiteralPath (Join-Path $payload "JSC-Licenses/$name") -Raw
    if (-not $notice.Contains('JSC-SOURCE.txt') -or $notice.Contains('JSC-CORRESPONDING-SOURCE.zip') -or $notice.Contains('accompanying source archives')) { throw "Missing or stale source reference: $name" }
}
$fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $payload 'quickxl.exe')).FileVersion
if ($fileVersion -ne $version) { throw "Internal and published versions differ: $fileVersion / $version" }
$exeBytes = [IO.File]::ReadAllBytes((Join-Path $payload 'quickxl.exe'))
$peOffset = [BitConverter]::ToInt32($exeBytes, 0x3c)
if ([BitConverter]::ToUInt16($exeBytes, $peOffset + 4) -ne 0x8664) { throw 'The application must be Windows x64.' }
if (Test-Path -LiteralPath (Join-Path $payload 'QuickExcel.Avalonia.exe')) { throw 'The old Windows apphost must not enter new packages.' }
foreach ($relative in @('QuickExcel.Avalonia.dll','QuickExcel.Avalonia.deps.json','QuickExcel.Avalonia.runtimeconfig.json','coreclr.dll','ScriptWorker/QuickExcel.ScriptWorker.exe','ScriptWorker/JavaScriptCore.dll','ScriptWorker/vcruntime140.dll','ScriptWorker/msvcp140.dll','Web/manifest.json','Licenses/CLICKONCE-LICENSE.txt','Licenses/WINDOWS-INSTALLER-NOTICES.txt')) {
    if (-not (Test-Path -LiteralPath (Join-Path $payload $relative))) { throw "Missing prerequisite or notice: $relative" }
}
& "$PSScriptRoot/test-windows-payload.ps1" -PackageDirectory $payload
$webview = @(Get-ChildItem -LiteralPath $package -Recurse -Filter MicrosoftEdgeWebview2Setup.exe)
if ($webview.Count -ne 1) { throw 'Exactly one WebView2 online bootstrapper is required.' }
Assert-QuickExcelWebView2Bootstrapper $webview[0].FullName
if (@(Get-ChildItem -LiteralPath $package -Recurse -Filter 'MicrosoftEdgeWebView2RuntimeInstaller*.exe').Count) { throw 'The large offline WebView2 installer must not enter the package.' }
# Run the exact ClickOnce payload, keeping all user data in fresh evidence.
$check = Join-Path $evidence 'self-check'
$process = Start-Process -FilePath (Join-Path $payload 'quickxl.exe') -ArgumentList @('--self-check',('"{0}"' -f $check)) -WorkingDirectory $payload -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $evidence 'stdout.log') -RedirectStandardError (Join-Path $evidence 'stderr.log')
try {
    if (-not $process.WaitForExit(90000)) { throw 'ClickOnce payload self-check timed out.' }
    $process.WaitForExit()
    $result = Get-Content -LiteralPath (Join-Path $check 'result.json') -Raw | ConvertFrom-Json
    if ($process.ExitCode -ne 0 -or -not $result.passed) { throw "ClickOnce payload self-check failed: $($result.error)" }
} finally { if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }; $process.Dispose() }
[ordered]@{ passed=$true; version=$version; manifestFileCount=$seen.Count; xmlSignaturesVerified=$true; fileHashesVerified=$true; codeSignaturesVerified=$true; jscSourceUrl=$sourceUrl; jscSourceArchiveIncluded=$false; certificateTrust='Self-signed; no trust store changes'; selfCheck=$result; installedThroughClickOnce=$false } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $evidence 'result.json') -Encoding utf8
Write-Host "PASS: ClickOnce signatures, $($seen.Count) file hashes, versions, prerequisites and production self-check."
