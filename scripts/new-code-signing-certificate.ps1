[CmdletBinding()]
param([string]$Subject = 'CN=QuickExcel', [string]$OutputDirectory = '')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (-not $IsWindows) { throw 'Code-signing certificates are created on Windows.' }
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts = Join-Path $root 'artifacts'
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $artifacts ('signing/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8)) }
$directory = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $directory.StartsWith($artifacts + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Certificate backups must stay under artifacts.' }
if (Test-Path -LiteralPath $directory) { throw 'Use a new certificate backup directory.' }
for ($parent = Split-Path $directory -Parent; $parent; $parent = Split-Path $parent -Parent) {
    if ((Test-Path -LiteralPath $parent) -and ((Get-Item -LiteralPath $parent).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Certificate paths cannot traverse reparse points.' }
}
New-Item -ItemType Directory -Path $directory | Out-Null
$acl = [Security.AccessControl.DirectorySecurity]::new()
$acl.SetAccessRuleProtection($true, $false)
$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
$acl.SetOwner($sid)
$acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
Set-Acl -LiteralPath $directory -AclObject $acl
$start = Get-Date
$certificate = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject -FriendlyName 'QuickExcel Code Signing (10 years)' -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotBefore $start -NotAfter $start.AddYears(10)
$password = ConvertTo-SecureString ([Convert]::ToBase64String([Security.Cryptography.RandomNumberGenerator]::GetBytes(48))) -AsPlainText -Force
Export-PfxCertificate -Cert $certificate -FilePath (Join-Path $directory 'QuickExcel-CodeSigning.pfx') -Password $password -CryptoAlgorithmOption AES256_SHA256 | Out-Null
Export-Certificate -Cert $certificate -FilePath (Join-Path $directory 'QuickExcel-CodeSigning.cer') | Out-Null
# The random PFX password never enters arguments, logs, source control or the installer.
$password | ConvertFrom-SecureString | Set-Content -LiteralPath (Join-Path $directory 'pfx-password.dpapi') -Encoding ascii -NoNewline
$metadata = [ordered]@{ subject=$certificate.Subject; thumbprint=$certificate.Thumbprint; notBefore=$certificate.NotBefore.ToString('o'); notAfter=$certificate.NotAfter.ToString('o'); store='Cert:\CurrentUser\My'; pfxPasswordProtection='Windows DPAPI, current user and machine'; directory=$directory }
$metadata | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'certificate.json') -Encoding utf8
@'
This directory is a PRIVATE signing backup. Do not distribute the PFX or DPAPI file.
Only QuickExcel-CodeSigning.cer is public and may accompany the installer.
The key is also installed in Cert:\CurrentUser\My (not in trusted roots).
To recover the PFX password as a SecureString on this Windows account/machine:
  $password = Get-Content ./pfx-password.dpapi | ConvertTo-SecureString
To make a portable backup with your own password (PowerShell):
  $cert = Get-Item ('Cert:\CurrentUser\My\' + (Get-Content ./certificate.json -Raw | ConvertFrom-Json).thumbprint)
  Export-PfxCertificate -Cert $cert -FilePath ./QuickExcel-Portable.pfx -Password (Read-Host 'New backup password' -AsSecureString)
Keep the same certificate for future ClickOnce updates. Back it up before migrating machines.
'@ | Set-Content -LiteralPath (Join-Path $directory 'README.txt') -Encoding utf8
$metadata | ConvertTo-Json
