[CmdletBinding()]
param([ValidateSet('linux-x64','osx-arm64')][string]$Rid='linux-x64', [switch]$CreateArchive)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$version=([xml](Get-Content -LiteralPath (Join-Path $root 'QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj') -Raw)).SelectSingleNode('/Project/PropertyGroup/Version').InnerText
$id=[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8)
$run=Join-Path $root "artifacts/avalonia/platform-releases/$Rid-$id"
$package=Join-Path $run 'application';$locks=Join-Path $run 'locks';$build=Join-Path $run 'build'
$native=Join-Path $root "artifacts/native/$Rid"
$library=if($Rid -eq 'linux-x64'){'libquickexcel-js.so'}else{'libquickexcel-js.dylib'}
foreach($name in @($library,'build-identity.txt')+$(if($Rid -eq 'linux-x64'){@('quickexcel-sandbox','libquickexcel-sandbox.so','sandbox-build-identity.txt')}else{@()})) {
    if(-not (Test-Path -LiteralPath (Join-Path $native $name) -PathType Leaf)){throw "Missing native input: $native/$name. Build on the target OS first."}
}
New-Item -ItemType Directory -Path $package,$locks | Out-Null
Push-Location $root
try {
    & "$PSScriptRoot/build-avalonia-web.ps1"
    foreach($lock in @(& git ls-files '*packages.lock.json')) {
        $projects=@(Get-ChildItem -LiteralPath (Split-Path (Join-Path $root $lock) -Parent) -Filter '*.csproj')
        if($projects.Count -eq 1){Copy-Item -LiteralPath (Join-Path $root $lock) -Destination (Join-Path $locks ($projects[0].BaseName+'.packages.lock.json'))}
    }
    $properties=@("-p:RuntimeIdentifier=$Rid","-p:QuickExcelNativeRid=$Rid","-p:QuickExcelReleaseLockDirectory=$locks",'-p:SelfContained=true','-p:QuickExcelDiagnostics=false','-p:DebugType=None','-p:DebugSymbols=false')
    dotnet restore QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -r $Rid --artifacts-path $build @properties -p:RestoreForceEvaluate=true -v minimal
    if($LASTEXITCODE -ne 0){throw 'Target RID restore failed.'}
    dotnet publish QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release -r $Rid --self-contained true --artifacts-path $build @properties -p:RestoreLockedMode=true -o $package -v minimal 2>&1 | Tee-Object (Join-Path $run 'publish.log')
    if($LASTEXITCODE -ne 0){throw 'Target RID publish failed.'}
    $all=@(Get-ChildItem -LiteralPath $package -File -Recurse)
    if(@($all | Where-Object {$_.Name -match 'WorkerFixture|HostFixture|SandboxProbe|PlatformSmoke|TestProbe|Designer|Avalonia\.Diagnostics|\.pdb$|\.map$|\.pfx$|ExcelIO|Grid\.Wpf|^PresentationFramework|^PresentationCore|Xpf|quickjs-eval|^quickexcel-js\.dll$'})){throw 'Forbidden platform payload.'}
    foreach($file in @('QuickExcel.Avalonia','QuickExcel.Avalonia.deps.json','Web/manifest.json',$library)+$(if($Rid -eq 'linux-x64'){@('libcoreclr.so','ScriptWorker/QuickExcel.ScriptWorker','ScriptWorker/libcoreclr.so','ScriptWorker/libquickexcel-js.so','ScriptWorker/quickexcel-sandbox','ScriptWorker/libquickexcel-sandbox.so')}else{@('libcoreclr.dylib')})) {
        if(-not(Test-Path -LiteralPath (Join-Path $package $file) -PathType Leaf)){throw "Missing payload: $file"}
    }
    if($Rid -eq 'osx-arm64' -and (Test-Path (Join-Path $package 'ScriptWorker'))){throw 'Unqualified macOS worker must not be published.'}
    if(@($all | Where-Object {$_.FullName.Contains([IO.Path]::DirectorySeparatorChar+'ScriptWorker'+[IO.Path]::DirectorySeparatorChar) -and $_.Name -match 'Avalonia|WebView|Skia|HarfBuzz|RichTextKit|PdfPig|Rendering|LibTiff|Monaco'})){throw 'UI payload in worker.'}
    $web=Get-Content (Join-Path $package 'Web/manifest.json') -Raw | ConvertFrom-Json
    if($web.testBuild -or $web.probeBuild){throw 'Production requires production Web resources.'}
    foreach($asset in $web.files) { $path=Join-Path $package ('Web/'+$asset.name);if((Get-Item $path).Length -ne $asset.size -or (Get-FileHash $path -Algorithm SHA256).Hash -ne $asset.sha256){throw 'Web manifest mismatch.'} }
    if(@(Get-ChildItem (Join-Path $package 'Web') -File -Recurse).Count -ne $web.files.Count+1){throw 'Unlisted Web assets.'}
    $metadata=Join-Path $package 'metadata';New-Item -ItemType Directory -Path $metadata | Out-Null
    Copy-Item -LiteralPath $locks -Destination $metadata -Recurse
    Copy-Item -LiteralPath (Join-Path $root 'QuickExcel.Web/package-lock.json') -Destination $metadata
    Copy-Item -Path (Join-Path $native '*identity.txt') -Destination $metadata
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'README-avalonia-unix.txt') -Destination (Join-Path $package 'README.txt')
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'start-quickexcel-unix.sh') -Destination (Join-Path $package 'Start-QuickExcel.sh')
    $files=@(Get-ChildItem -LiteralPath $package -File -Recurse | Sort-Object FullName | ForEach-Object {[ordered]@{path=[IO.Path]::GetRelativePath($package,$_.FullName).Replace('\','/');size=$_.Length;sha256=(Get-FileHash $_.FullName -Algorithm SHA256).Hash}})
    [ordered]@{schemaVersion=1;product='QuickExcel Avalonia';version=$version;rid=$Rid;selfContained=$true;diagnostics=$false;sourceHead=(& git rev-parse HEAD);workingTreeDirty=[bool](& git status --porcelain);sdk=(& dotnet --version);scriptCapability=$(if($Rid -eq 'linux-x64'){'Landlock/seccomp/cgroup: checked on every launch'}else{'disabled: macOS sandbox not qualified'});acceptance='Target-device acceptance deferred; package requires target self-check before distribution';webPackages=$web.packages;files=$files} | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $package 'release.json') -Encoding utf8
    if($CreateArchive) {
        # Python sets POSIX executable modes explicitly even for a Windows cross-publish.
        & python "$PSScriptRoot/archive-unix-package.py" $package (Join-Path $run "QuickExcel-$Rid.tar.gz")
        if($LASTEXITCODE -ne 0){throw 'Archive verification failed.'}
    }
    Write-Host "PASS: $package (target self-check and device gates are separate)"
} finally {Pop-Location}

