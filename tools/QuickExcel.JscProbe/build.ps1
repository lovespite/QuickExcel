param([ValidateRange(1, 16)][int]$Parallel = 6, [switch]$PrepareOnly, [switch]$IncludeProbe)
# dotnet launched from PowerShell 7 can pass a module path incompatible with 5.1.
if ($PSVersionTable.PSVersion.Major -eq 5) {
    $env:PSModulePath = (Join-Path $PSHOME 'Modules') + [IO.Path]::PathSeparator + $env:PSModulePath
    Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -ErrorAction Stop
}
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$root = Join-Path $repo 'artifacts\jsc'
& python (Join-Path $PSScriptRoot 'prepare.py')
if ($LASTEXITCODE -ne 0) { throw 'J0 input verification failed.' }
& python (Join-Path $PSScriptRoot 'prepare_compiler.py')
if ($LASTEXITCODE -ne 0) { throw 'TypeScript payload preparation failed.' }
if ($PrepareOnly) { return }
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
# Hosted runners service compiler/redist patch versions independently and may
# install a newer VS alongside the v143 toolset. Select an installation that
# actually contains our required 14.44 toolset rather than the newest VS alone.
$vs = $null
$vcToolsVersion = $null
$crt = $null
$crtNames = @('msvcp140.dll','msvcp140_2.dll','vcruntime140.dll','vcruntime140_1.dll')
$installations = @(& $vswhere -all -products * -sort -property installationPath)
if ($LASTEXITCODE -ne 0) { throw 'Visual Studio installation discovery failed.' }
foreach ($installation in $installations) {
    $toolsets = @(Get-ChildItem -LiteralPath (Join-Path $installation 'VC/Tools/MSVC') -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^14\.44\.\d+$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'bin/Hostx64/x64/cl.exe')) } |
        Sort-Object { [version]$_.Name } -Descending)
    $redistributables = @(Get-ChildItem -LiteralPath (Join-Path $installation 'VC/Redist/MSVC') -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^14\.44\.\d+$' } |
        Sort-Object { [version]$_.Name } -Descending)
    foreach ($redistributable in $redistributables) {
        $candidateCrt = Join-Path $redistributable.FullName 'x64/Microsoft.VC143.CRT'
        if ($toolsets.Count -gt 0 -and @($crtNames | Where-Object { -not (Test-Path -LiteralPath (Join-Path $candidateCrt $_) -PathType Leaf) }).Count -eq 0) {
            $vs = $installation
            $vcToolsVersion = $toolsets[0].Name
            $crt = $candidateCrt
            break
        }
    }
    if ($vs) { break }
}
if (-not $vs) { throw 'Visual Studio with the MSVC v143 14.44 x64 tools and matching 14.44 x64 redistributables is required.' }
Import-Module (Join-Path $vs 'Common7\Tools\Microsoft.VisualStudio.DevShell.dll')
Enter-VsDevShell -VsInstallPath $vs -SkipAutomaticLocation -DevCmdArguments "-arch=x64 -host_arch=x64 -vcvars_ver=$vcToolsVersion"
if ($env:VCToolsVersion.TrimEnd('\','/') -ne $vcToolsVersion) { throw 'Visual Studio did not activate the selected MSVC 14.44 toolset.' }
$env:OS = 'Windows_NT'
$env:PROCESSOR_ARCHITECTURE = 'AMD64'
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
foreach ($name in @('MSBuildExtensionsPath','MSBuildExtensionsPath32','MSBuildExtensionsPath64','MSBuildSDKsPath','MSBUILD_EXE_PATH','MSBuildToolsPath')) {
    Remove-Item "Env:$name" -ErrorAction SilentlyContinue
}
$llvm = Join-Path $root 'tools\llvm\bin'
$ruby = Join-Path $root 'tools\ruby\rubyinstaller-4.0.9-1-x64\bin\ruby.exe'
$perl = Join-Path $env:ProgramFiles 'Git\usr\bin\perl.exe'
$ninja = Join-Path $vs 'Common7\IDE\CommonExtensions\Microsoft\CMake\Ninja\ninja.exe'
$env:Path = "$llvm;$(Split-Path -Parent $perl);$env:Path"
$runId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8)
$logs = Join-Path $root "logs\$runId"
New-Item -ItemType Directory -Force $logs | Out-Null
Write-Host "Build logs: $logs"
$icu = Join-Path $root 'source\icu'
# ICU solution targets share intermediate directories. Build them serially to
# prevent duplicate dependency nodes racing over tlogs/data; JSC still uses Parallel.
# makedata builds upstream tool/test dependencies; it does not run regression tests.
& (Join-Path $vs 'MSBuild\Current\Bin\amd64\MSBuild.exe') (Join-Path $icu 'source\allinone\allinone.sln') /m:1 /p:Configuration=Release /p:Platform=x64 /p:PlatformToolset=v143 "/p:VCToolsVersion=$vcToolsVersion" '/t:common;i18n;makedata' /v:minimal /nologo *> (Join-Path $logs 'icu-build.log')
if ($LASTEXITCODE -ne 0) { throw "ICU build failed; see $logs\icu-build.log" }
if ((Get-Item (Join-Path $icu 'bin64\icudt77.dll')).Length -lt 1MB) { throw 'ICU data DLL is still the empty build stub.' }
$build = Join-Path $root 'build\win-x64'
$arguments = @('-S', (Join-Path $root 'source\webkitgtk-2.54.2'), '-B', $build, '-G', 'Ninja',
    "-DCMAKE_MAKE_PROGRAM=$ninja", "-DCMAKE_C_COMPILER=$llvm\clang-cl.exe", "-DCMAKE_CXX_COMPILER=$llvm\clang-cl.exe",
    '-DCMAKE_BUILD_TYPE=Release', '-DPORT=JSCOnly', '-DENABLE_FTL_JIT=OFF', '-DENABLE_WEBASSEMBLY=OFF',
    '-DUSE_SYSTEM_MALLOC=OFF', '-DUSE_MIMALLOC=ON',
    '-DENABLE_REMOTE_INSPECTOR=OFF', '-DENABLE_API_TESTS=OFF', '-DDEVELOPER_MODE=OFF',
    "-DICU_ROOT=$icu", "-DPERL_EXECUTABLE=$perl", "-DRuby_EXECUTABLE=$ruby",
    "-DPython_EXECUTABLE=$((Get-Command python).Source)",
    "-DQX_JSC_PROBE_SOURCE=$($PSScriptRoot.Replace('\','/'))",
    "-DCMAKE_PROJECT_WebKit_INCLUDE=$($PSScriptRoot.Replace('\','/'))/cmake-hook.cmake")
& cmake @arguments *> (Join-Path $logs 'configure.log')
if ($LASTEXITCODE -ne 0) { throw "JSC configure failed; see $logs\configure.log" }
& cmake --build $build --target $(if ($IncludeProbe) { 'quickexcel-jsc-probe' } else { 'JavaScriptCore' }) --parallel $Parallel *> (Join-Path $logs 'build.log')
if ($LASTEXITCODE -ne 0) { throw "JSC build failed; see $logs\build.log" }
foreach ($name in @('icudt77.dll','icuin77.dll','icuuc77.dll')) {
    Copy-Item -LiteralPath (Join-Path $icu "bin64\$name") -Destination (Join-Path $build 'bin') -Force
}
foreach ($name in $crtNames) {
    Copy-Item -LiteralPath (Join-Path $crt $name) -Destination (Join-Path $build 'bin') -Force
}
$identity = [ordered]@{
    status = 'jsc-production-native'; openIssues = @('JSC-MEM-001'); utc = [DateTime]::UtcNow.ToString('o')
    sourceManifest = (Get-FileHash (Join-Path $PSScriptRoot 'sources.json')).Hash
    windowsSourceManifest = (Get-FileHash (Join-Path $PSScriptRoot 'windows-sources.json')).Hash
    sourceAdjustments = (Get-FileHash (Join-Path $PSScriptRoot 'source-adjustments.json')).Hash
    cmake = (& cmake --version | Select-Object -First 1); compiler = (& "$llvm\clang-cl.exe" --version | Select-Object -First 1)
    compilerSha256 = (Get-FileHash "$llvm\clang-cl.exe").Hash
    visualStudio = $vs; vcTools = $env:VCToolsVersion; vcRedist = $crt; windowsSdk = $env:WindowsSDKVersion
    logs = $logs
    arguments = $arguments
    probeSources = @('probe.cpp','CMakeLists.txt','cmake-hook.cmake','build.ps1','../../QuickExcel.Scripting/native/jsc/bridge.cpp') | ForEach-Object {
        @{name=$_;sha256=(Get-FileHash (Join-Path $PSScriptRoot $_)).Hash}
    }
    binaries = (@('JavaScriptCore.dll','icudt77.dll','icuin77.dll','icuuc77.dll','msvcp140.dll','msvcp140_2.dll','vcruntime140.dll','vcruntime140_1.dll') + $(if ($IncludeProbe) { @('quickexcel-jsc-probe.exe') } else { @() })) | ForEach-Object {
        @{name=$_;sha256=(Get-FileHash (Join-Path $build "bin\$_")).Hash}
    }
}
$identity | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $build 'identity.json') -Encoding utf8
& python (Join-Path $PSScriptRoot 'prepare_distribution.py')
if ($LASTEXITCODE -ne 0) { throw 'JSC distribution materials failed.' }
Write-Host "JSC native payload: $build\bin; development probe included: $IncludeProbe"
