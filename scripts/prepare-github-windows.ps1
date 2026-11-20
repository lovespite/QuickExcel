[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (-not $IsWindows -or $env:GITHUB_ACTIONS -ne 'true') { throw 'This prerequisite installer is only for a Windows GitHub Actions runner.' }
. "$PSScriptRoot/windows-signing.ps1"
function Test-WebViewRuntime {
    $key = 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
        $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, [Microsoft.Win32.RegistryView]::Registry32)
        $registration = $null; $state = $null
        try {
            $registration = $registry.OpenSubKey($key)
            $state = $registry.OpenSubKey($key.Replace('\Clients\', '\ClientState\'))
            $version = if ($registration) { $registration.GetValue('pv') } else { $null }
            $directory = if ($state) { $state.GetValue('EBWebView') } else { $null }
            $parsed = $null
            if ($version -and [version]::TryParse($version, [ref]$parsed) -and $parsed.Major -gt 0 -and $directory -and
                [IO.Path]::IsPathFullyQualified($directory) -and
                (Test-Path -LiteralPath (Join-Path $directory 'msedgewebview2.exe')) -and
                (Test-Path -LiteralPath (Join-Path $directory 'EBWebView/x64/EmbeddedBrowserWebView.dll'))) { return $true }
        } finally {
            if ($registration) { $registration.Dispose() }
            if ($state) { $state.Dispose() }
            $registry.Dispose()
        }
    }
    return $false
}
if (Test-WebViewRuntime) { Write-Host 'WebView2 Evergreen Runtime is installed.'; return }
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$downloads = Join-Path $root 'artifacts/github-prerequisites'
New-Item -ItemType Directory -Path $downloads -Force | Out-Null
$installer = Join-Path $downloads 'MicrosoftEdgeWebview2Setup.exe'
Invoke-WebRequest 'https://go.microsoft.com/fwlink/p/?LinkId=2124703' -OutFile $installer
Assert-QuickExcelWebView2Bootstrapper $installer
$process = Start-Process -FilePath $installer -ArgumentList @('/silent', '/install') -WindowStyle Hidden -PassThru
try {
    if (-not $process.WaitForExit(300000)) { throw 'WebView2 installation exceeded five minutes.' }
    $process.WaitForExit()
    if ($process.ExitCode -notin @(0, 3010)) { throw "WebView2 installer failed with exit code $($process.ExitCode)." }
} finally {
    if (-not $process.HasExited) { $process.Kill($true); [void]$process.WaitForExit(5000) }
    $process.Dispose()
}
if (-not (Test-WebViewRuntime)) { throw 'WebView2 installation did not register a usable x64 runtime.' }
Write-Host 'PASS: WebView2 Evergreen Runtime is ready for production self-check.'
