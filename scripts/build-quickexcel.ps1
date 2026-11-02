[CmdletBinding()]
param([ValidateSet('Release')][string]$Configuration='Release', [switch]$SkipTests,
      [string]$OutputDirectory='', [switch]$RunSmoke, [switch]$CreateArchive)
# Default delivery is the self-contained Avalonia Windows package. Activation always
# requires its production self-check; individual feature work uses targeted scripts.
& "$PSScriptRoot/build-avalonia-release.ps1" -SkipTests:$SkipTests -OutputDirectory $OutputDirectory -RunSmoke:$RunSmoke -CreateArchive:$CreateArchive -Activate
if (-not $?) { throw 'QuickExcel release failed.' }
