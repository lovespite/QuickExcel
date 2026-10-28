[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PackageDirectory,
    [ValidateSet('portable-unicode-spaces','owner-read-execute','read-only-file-attributes','long-path','long-path-dotnet','replace-and-uninstall')]
    [string[]]$CasesToRun = @('portable-unicode-spaces','owner-read-execute','read-only-file-attributes','long-path','long-path-dotnet','replace-and-uninstall')
)

# Local deployment evidence only. This does not certify a clean machine or another Windows account.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repo = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$artifacts = Join-Path $repo 'artifacts'
$source = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($PackageDirectory)) { $PackageDirectory } else { Join-Path $repo $PackageDirectory }))
if (-not $source.StartsWith($artifacts + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'PackageDirectory must be inside this repository artifacts directory.' }
$run = Join-Path $artifacts ('deployment-matrix\' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $run | Out-Null

function Assert-NoLinks([string]$Path) {
    $parent = $Path
    while ($parent) {
        if (Test-Path -LiteralPath $parent) {
            if (((Get-Item -LiteralPath $parent -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Linked path refused: $parent" }
        }
        $parent = Split-Path $parent -Parent
    }
    foreach ($item in Get-ChildItem -LiteralPath $Path -Recurse -Force) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Linked content refused: $($item.FullName)" }
    }
}
function Remove-TestTree([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($run + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing cleanup outside this run: $full" }
    if (Test-Path -LiteralPath $full) { Assert-NoLinks $full; Remove-Item -LiteralPath $full -Recurse -Force }
}
function Copy-Package([string]$Destination) {
    New-Item -ItemType Directory -Path $Destination | Out-Null
    foreach ($item in Get-ChildItem -LiteralPath $source -Force) { Copy-Item -LiteralPath $item.FullName -Destination $Destination -Recurse }
}
function Get-PackageState([string]$Directory) {
    $items = @((Get-Item -LiteralPath $Directory)) + @(Get-ChildItem -LiteralPath $Directory -Recurse -Force | Sort-Object FullName)
    return @($items | ForEach-Object {
        [ordered]@{ path = $_.FullName.Substring($Directory.Length); acl = (Get-Acl -LiteralPath $_.FullName).Sddl;
            attributes = [string]$_.Attributes; sha256 = $(if ($_.PSIsContainer) { $null } else { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) }
    })
}
function Run-Smoke([string]$Directory, [string]$Label, [bool]$ManagedLaunch = $false) {
    $evidence = Join-Path $run ('evidence\' + $Label)
    New-Item -ItemType Directory -Path $evidence | Out-Null
    $before = Get-PackageState $Directory | ConvertTo-Json -Depth 5 -Compress
    $start = [Diagnostics.ProcessStartInfo]::new($(if ($ManagedLaunch) { (Get-Command dotnet -CommandType Application).Source } else { Join-Path $Directory 'QuickExcel.exe' }))
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true; $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    # CreateProcess has a MAX_PATH limit for lpCurrentDirectory; application paths are independent of CWD.
    $start.WorkingDirectory = $run
    if ($ManagedLaunch) { $start.ArgumentList.Add((Join-Path $Directory 'QuickExcel.dll')) }
    $start.ArgumentList.Add('--smoke-test'); $start.ArgumentList.Add($evidence)
    $process = [Diagnostics.Process]::Start($start)
    try {
        if (-not $process.WaitForExit(45000)) { throw 'Desktop smoke exceeded 45 seconds.' }
        $process.WaitForExit(); $process.Refresh()
        if ($process.ExitCode -ne 0) { throw "Desktop smoke exit $($process.ExitCode); see $evidence" }
        foreach ($name in @('smoke-result.txt','storage-smoke-result.txt','close-smoke-result.txt','calculation-smoke-result.txt','date-display-smoke-result.txt','worker-smoke-result.txt','actions-smoke-result.txt')) {
            $result = Join-Path $evidence $name
            if (-not (Test-Path -LiteralPath $result) -or -not (Get-Content -LiteralPath $result -Raw).StartsWith('PASS:', [StringComparison]::Ordinal)) { throw "Missing PASS: $name" }
        }
        $after = Get-PackageState $Directory | ConvertTo-Json -Depth 5 -Compress
        if ($before -cne $after) { throw 'Smoke changed installation contents, file attributes or ACLs.' }
    }
    finally {
        if (-not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
        $process.Dispose()
    }
    return $evidence
}

Assert-NoLinks $source
Assert-NoLinks $run
foreach ($required in @('QuickExcel.exe','BUILD-METADATA.json','ScriptWorker\QuickExcel.ScriptWorker.exe','ScriptWorker\quickexcel-js.dll')) {
    if (-not (Test-Path -LiteralPath (Join-Path $source $required))) { throw "Incomplete package: $required" }
}
if (Get-ChildItem -LiteralPath $source -Recurse -File | Where-Object { $_.Name -match 'Fixture|SandboxProbe' }) { throw 'Development-only binary found in package.' }
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$owner = $identity.User
$longPathsEnabled = Get-ItemPropertyValue -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name LongPathsEnabled -ErrorAction SilentlyContinue
$report = [ordered]@{
    utc = [DateTime]::UtcNow.ToString('o'); os = [Environment]::OSVersion.VersionString; userSid = $owner.Value;
    elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);
    longPathsEnabled = $longPathsEnabled;
    source = $source; buildMetadataSha256 = (Get-FileHash -LiteralPath (Join-Path $source 'BUILD-METADATA.json')).Hash;
    cleanMachineValidated = $false; launch = 'CreateProcess with short working directory'; cases = @();
    pending = @('Separate standard-user account', 'Elevated administrator account', 'Actual Program Files installation', 'Explorer / cmd long-path launch', 'Cross-version upgrade', 'Clean machine / missing runtime', 'M6 QA-OS-01')
}
$longPath = 'long-path'
while ((Join-Path $run $longPath).Length -lt 275) { $longPath = Join-Path $longPath 'long-folder-012345678901234567890123456789' }
$cases = @(
    @{ name = 'portable-unicode-spaces'; path = '便携 package with spaces' },
    @{ name = 'owner-read-execute'; path = 'read-only-acl' },
    @{ name = 'read-only-file-attributes'; path = 'read-only-attributes' },
    @{ name = 'long-path'; path = $longPath },
    @{ name = 'long-path-dotnet'; path = (Join-Path $longPath 'managed-launch') },
    @{ name = 'replace-and-uninstall'; path = 'replace-installation' }
)
foreach ($case in $cases) {
    if ($case.name -notin $CasesToRun) { continue }
    $directory = Join-Path $run $case.path
    $row = [ordered]@{ name = $case.name; installation = $directory; executablePathLength = (Join-Path $directory 'QuickExcel.exe').Length; status = 'failed'; evidence = @(); error = $null }
    $savedAcls = @()
    try {
        Copy-Package $directory
        if ($case.name -eq 'owner-read-execute') {
            $items = @((Get-Item -LiteralPath $directory)) + @(Get-ChildItem -LiteralPath $directory -Recurse -Force)
            foreach ($item in $items) {
                $savedAcls += @{ path = $item.FullName; sddl = (Get-Acl -LiteralPath $item.FullName).Sddl; directory = $item.PSIsContainer }
                $acl = if ($item.PSIsContainer) { [Security.AccessControl.DirectorySecurity]::new() } else { [Security.AccessControl.FileSecurity]::new() }
                $acl.SetAccessRuleProtection($true, $false)
                $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($owner, [Security.AccessControl.FileSystemRights]::ReadAndExecute, [Security.AccessControl.AccessControlType]::Allow))
                [IO.FileSystemAclExtensions]::SetAccessControl($item, $acl)
            }
            $denied = $false
            try { [IO.File]::WriteAllText((Join-Path $directory 'write-probe'), 'must be denied') }
            catch [UnauthorizedAccessException] { $denied = $true }
            if (-not $denied) { throw 'Read-only ACL positive check failed: installation was writable.' }
        }
        if ($case.name -eq 'read-only-file-attributes') {
            Get-ChildItem -LiteralPath $directory -Recurse -File | ForEach-Object { $_.Attributes = $_.Attributes -bor [IO.FileAttributes]::ReadOnly }
        }
        $row.evidence += Run-Smoke $directory $case.name ($case.name -eq 'long-path-dotnet')
        if ($case.name -eq 'replace-and-uninstall') {
            Remove-TestTree $directory
            Copy-Package $directory
            $row.evidence += Run-Smoke $directory ($case.name + '-reinstalled')
        }
        $row.status = 'passed'
    }
    catch { $row.error = $_.Exception.ToString() }
    finally {
        try {
            foreach ($saved in $savedAcls) {
                $restoreAcl = if ($saved.directory) { [Security.AccessControl.DirectorySecurity]::new() } else { [Security.AccessControl.FileSecurity]::new() }
                $restoreAcl.SetSecurityDescriptorSddlForm($saved.sddl, [Security.AccessControl.AccessControlSections]::Access)
                [IO.FileSystemAclExtensions]::SetAccessControl((Get-Item -LiteralPath $saved.path -Force), $restoreAcl)
            }
            Remove-TestTree $directory
            $row.uninstallSucceeded = -not (Test-Path -LiteralPath $directory)
        }
        catch { $row.status = 'failed'; $row.error = [string]$row.error + "`nCleanup: " + $_.Exception.ToString() }
        $report.cases += $row
        $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run 'matrix.json') -Encoding UTF8
        Write-Host "$($case.name): $($row.status)"
    }
}
Write-Host "Evidence: $run"
if (@($report.cases | Where-Object status -ne 'passed').Count -ne 0) { throw 'Deployment matrix has failed cases; inspect matrix.json.' }
