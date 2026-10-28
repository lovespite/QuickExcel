[CmdletBinding()]
param([Parameter(Mandatory)][string]$EvidenceDirectory)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = [IO.Path]::GetFullPath((Join-Path $repository 'artifacts\sandbox-probe')) + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceDirectory)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence must stay under repository artifacts/sandbox-probe.' }
$request = Join-Path $evidence 'capture-request.json'
$deadline = [DateTime]::UtcNow.AddSeconds(90)
while (-not (Test-Path -LiteralPath $request)) {
    if ([DateTime]::UtcNow -gt $deadline) { throw 'Timed out waiting for probe.' }
    Start-Sleep -Milliseconds 200
}
try {
    $config = Get-Content -LiteralPath $request -Raw | ConvertFrom-Json
    $executable = [IO.Path]::GetFullPath([string]$config.executable)
    if ([IO.Path]::GetFileName($executable) -ne 'QuickExcel.SandboxProbe.exe' -or -not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        throw 'Expected the running development probe payload.'
    }
    # Read-only, filtered to this unique executable. No firewall rules, buffering settings or loopback exemptions are modified.
    & netsh wfp show options optionsfor=NETEVENTS | Set-Content -LiteralPath (Join-Path $evidence 'wfp-options.txt')
    & netsh wfp show netevents "file=$(Join-Path $evidence 'wfp-netevents.xml')" "appid=$executable" |
        Set-Content -LiteralPath (Join-Path $evidence 'wfp-events-command.txt')
    if ($LASTEXITCODE -ne 0) { throw 'WFP event query failed.' }
    # Loopback drops occur on the RECEIVING listener, whose appId is the parent, not the isolated child.
    foreach ($version in @('IPv4', 'IPv6')) {
        $port = [int]$config.($version + 'Port')
        if ($port -lt 1 -or $port -gt 65535) { throw 'Invalid listener port.' }
        & netsh wfp show netevents "file=$(Join-Path $evidence "wfp-loopback-$version.xml")" "localport=$port" protocol=6 |
            Set-Content -LiteralPath (Join-Path $evidence "wfp-loopback-$version-command.txt")
        if ($LASTEXITCODE -ne 0) { throw 'WFP loopback event query failed.' }
    }
    & netsh wfp show filters "file=$(Join-Path $evidence 'wfp-filters.xml')" "appid=$executable" verbose=ON |
        Set-Content -LiteralPath (Join-Path $evidence 'wfp-filters-command.txt')
    if ($LASTEXITCODE -ne 0) { throw 'WFP filter query failed.' }
    'Read-only WFP collection completed.' | Set-Content -LiteralPath (Join-Path $evidence 'wfp-result.txt')
}
catch {
    $_.ToString() | Set-Content -LiteralPath (Join-Path $evidence 'wfp-result.txt')
    exit 1
}
