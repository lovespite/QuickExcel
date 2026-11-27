[CmdletBinding()]
param([string]$Executable = 'artifacts/build/bin/QuickExcel.Desktop.Avalonia/release/quickxl.exe')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$run = Join-Path $root ('artifacts/avalonia/folder-picker/' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run | Out-Null
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type -Name FolderPickerWindow -Namespace QuickExcel.Diagnostics -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern bool PostMessage(System.IntPtr window, uint message, System.IntPtr wParam, System.IntPtr lParam);
'@
$element = [Windows.Automation.AutomationElement]
$scope = [Windows.Automation.TreeScope]
$process = Start-Process -FilePath (Resolve-Path $Executable) -WorkingDirectory $root -WindowStyle Hidden -PassThru -ArgumentList @('--web-smoke', ('"{0}"' -f $run), '--file-picker-only')
try {
    $own = New-Object Windows.Automation.PropertyCondition($element::ProcessIdProperty, $process.Id)
    $shellDialog = New-Object Windows.Automation.PropertyCondition($element::ClassNameProperty, '#32770')
    $until = [DateTime]::UtcNow.AddSeconds(45)
    do {
        $windows = $element::RootElement.FindAll($scope::Children, $own)
        foreach ($window in $windows) {
            $dialog = $window.FindAll($scope::Subtree, $shellDialog) |
                Where-Object { $_.Current.Name -like '*QuickExcel picker check*' } | Select-Object -First 1
            if ($dialog) { break }
        }
        if ($null -eq $dialog) { Start-Sleep -Milliseconds 200 }
    } while ($null -eq $dialog -and !$process.HasExited -and [DateTime]::UtcNow -lt $until)
    if ($null -eq $dialog) { throw "Missing native folder picker. Process exited: $($process.HasExited). Windows: $($windows | ForEach-Object { $_.Current.Name }). Evidence: $run" }
    $descendants = $dialog.FindAll($scope::Descendants, [Windows.Automation.Condition]::TrueCondition)
    $classes = @($descendants | ForEach-Object { $_.Current.ClassName } | Where-Object { $_ } | Sort-Object -Unique)
    [ordered]@{ title = $dialog.Current.Name; windowClass = $dialog.Current.ClassName; controls = $classes } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'dialog.json') -Encoding utf8
    if ('Breadcrumb Parent' -notin $classes -or 'SearchEditBoxWrapperClass' -notin $classes) {
        throw "Missing Shell address bar or search field. Evidence: $run"
    }
    if (![QuickExcel.Diagnostics.FolderPickerWindow]::PostMessage([IntPtr]$dialog.Current.NativeWindowHandle, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)) {
        throw "Could not dismiss the selected test dialog."
    }
    if (!$process.WaitForExit(30000)) { throw "Action did not resume. Evidence: $run" }
    $result = Get-Content -LiteralPath (Join-Path $run 'result.json') -Raw | ConvertFrom-Json
    if (!$result.passed -or $process.ExitCode -ne 0) { throw "Picker smoke failed. $($result.error) Evidence: $run" }
    Write-Host "PASS: native Explorer folder picker, address bar, search and cancellation. Evidence: $run"
} finally {
    if (!$process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
    $process.Dispose()
}
