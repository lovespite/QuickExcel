[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent)); $artifacts=Join-Path $root 'artifacts'
$source=[IO.Path]::GetFullPath($PackageDirectory)
if (-not $source.StartsWith($artifacts+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Package must be inside artifacts.' }
if (@(Get-ChildItem -LiteralPath $source -Recurse -Force | Where-Object {$_.Attributes -band [IO.FileAttributes]::ReparsePoint})) { throw 'Linked package refused.' }
$run=Join-Path $artifacts ('avalonia/deployment/'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $run|Out-Null
$results=[Collections.Generic.List[object]]::new()
function Copy-Package([string]$name) {
    $destination=Join-Path $run $name; New-Item -ItemType Directory -Path $destination|Out-Null
    Get-ChildItem -LiteralPath $source -Force|ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination $destination -Recurse}
    return $destination
}
function Check-Package([string]$directory,[string]$name,[bool]$expected) {
    $evidence=Join-Path $run ('evidence-'+$name)
    $start=[Diagnostics.ProcessStartInfo]::new((Join-Path $directory 'quickxl.exe'))
    $start.UseShellExecute=$false; $start.CreateNoWindow=$true; $start.WindowStyle=[Diagnostics.ProcessWindowStyle]::Hidden; $start.WorkingDirectory=$run
    $start.ArgumentList.Add('--self-check'); $start.ArgumentList.Add($evidence)
    $process=[Diagnostics.Process]::Start($start)
    try {
        if (-not $process.WaitForExit(90000)) {throw "Self-check timeout: $name"}
        $process.WaitForExit(); $result=Get-Content -LiteralPath (Join-Path $evidence 'result.json') -Raw|ConvertFrom-Json
        if ($result.passed -ne $expected -or ($expected -and $process.ExitCode -ne 0) -or (-not $expected -and $process.ExitCode -eq 0)) {throw "Unexpected deployment result: $name $($result.error)"}
        if ($name -eq 'missing-resource' -and $result.error -notmatch 'script\.js') { throw 'Resource failure was not caused by the removed asset.' }
        if ($name -eq 'missing-worker-no-fallback' -and ($result.checks.Count -ne 4 -or $result.error -notmatch 'worker')) { throw 'Worker failure did not occur at the isolated-worker check.' }
        $results.Add([ordered]@{name=$name;expectedSuccess=$expected;exitCode=$process.ExitCode;result=$result})
    } finally {if(-not $process.HasExited){$process.Kill();[void]$process.WaitForExit(5000)};$process.Dispose()}
}
$portable=Copy-Package '便携 安装 目录'; Check-Package $portable 'unicode-spaces' $true
$readonly=Copy-Package 'read-only-installation'; $original=Get-Acl -LiteralPath $readonly
try {
    $acl=Get-Acl -LiteralPath $readonly
    $user=[Security.Principal.WindowsIdentity]::GetCurrent().User
    $rule=[Security.AccessControl.FileSystemAccessRule]::new($user,[Security.AccessControl.FileSystemRights]::Write,[Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit',[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Deny)
    $acl.AddAccessRule($rule); Set-Acl -LiteralPath $readonly -AclObject $acl
    Check-Package $readonly 'read-only-installation' $true
} finally {Set-Acl -LiteralPath $readonly -AclObject $original}
$missing=Copy-Package 'missing-resource'
$file=[IO.Path]::GetFullPath((Join-Path $missing 'Web/script.js'))
if(-not $file.StartsWith($run+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test mutation.'}
Move-Item -LiteralPath $file -Destination ($file+'.removed')
Check-Package $missing 'missing-resource' $false
$worker=Copy-Package 'missing-worker'
$file=[IO.Path]::GetFullPath((Join-Path $worker 'ScriptWorker/QuickExcel.ScriptWorker.exe'))
if(-not $file.StartsWith($run+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test mutation.'}
Move-Item -LiteralPath $file -Destination ($file+'.removed')
Check-Package $worker 'missing-worker-no-fallback' $false
[ordered]@{passed=$true;source=$source;cases=$results;scope='Local ordinary-user copies only; clean-machine, physical devices and actual Program Files remain deferred.'}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $run 'result.json') -Encoding utf8
Write-Host "PASS: deployment cases $run"
