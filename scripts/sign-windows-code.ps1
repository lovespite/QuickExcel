[CmdletBinding()]
param([Parameter(Mandatory)][string]$CertificateThumbprint, [Parameter(Mandatory)][string[]]$Path)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot/windows-signing.ps1"
$certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$CertificateThumbprint"
if (-not $certificate.HasPrivateKey -or $certificate.NotAfter -le (Get-Date) -or $certificate.NotBefore -gt (Get-Date)) { throw 'A currently valid signing certificate with a private key is required.' }
if ('1.3.6.1.5.5.7.3.3' -notin @($certificate.EnhancedKeyUsageList | ForEach-Object { $_.ObjectId })) { throw 'Certificate must permit code signing.' }
$kits = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10/bin'
$signTool = Get-ChildItem -LiteralPath $kits -Directory | Where-Object { $_.Name -match '^10\.0\.\d+\.0$' } | Sort-Object { [version]$_.Name } -Descending | ForEach-Object { Join-Path $_.FullName 'x64/signtool.exe' } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $signTool) { throw 'Install the Windows SDK x64 SignTool.' }
foreach ($file in $Path) {
    & $signTool sign /sha1 $CertificateThumbprint /s My /fd SHA256 /d QuickExcel $file
    if ($LASTEXITCODE -ne 0) { throw "Code signing failed: $file" }
    Assert-QuickExcelCodeSignature $file $CertificateThumbprint
}
