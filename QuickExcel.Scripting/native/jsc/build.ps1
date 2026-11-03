param([ValidateRange(1, 16)][int]$Parallel = 6)
$ErrorActionPreference = 'Stop'
& "$PSScriptRoot/../../../tools/QuickExcel.JscProbe/build.ps1" -Parallel $Parallel
if (-not $?) { throw 'Pinned JSC native build failed.' }
