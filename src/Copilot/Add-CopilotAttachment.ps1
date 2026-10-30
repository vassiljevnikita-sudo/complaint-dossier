<#
.SYNOPSIS
  Attaches local files to the open Copilot chat without user input.

.DESCRIPTION
  For each file: opens the "Add sources" menu, chooses "Upload images and
  files", writes the full path into the native Windows file dialog that belongs
  to the Copilot window, presses Open, and waits until the chat shows the file
  as an attachment (a "remove <file>" control) and no upload is in progress.
  Fails if a step is ambiguous: two candidate dialogs, two upload buttons,
  or an upload error message. Runs in Windows PowerShell 5.1. ASCII only.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string[]]$FilePath,
    [string]$WindowTitle = 'Microsoft 365 Copilot',
    [int]$TimeoutMilliseconds = 60000
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
if (-not ('ComplaintDossierUploadWin32' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ComplaintDossierUploadWin32 {
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr hWnd, uint command);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint message, UIntPtr wParam, string text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint message, UIntPtr wParam, IntPtr value, uint flags, uint timeout, out UIntPtr result);
}
"@
}

$WM_SETTEXT = 0x000C
$BM_CLICK = 0x00F5
$GW_OWNER = 4
$UploadMenuPattern = '(?i)^(bilder und dateien hochladen|upload images and files|upload from device)$'

function Get-TopLevelWindows {
    @([System.Windows.Automation.AutomationElement]::RootElement.FindAll(
            [System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition))
}

function Get-Descendants($Root) {
    @($Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition))
}

function Get-CopilotRoots {
    # Pop-up menus can be separate top-level windows of the Copilot process tree.
    @(Get-TopLevelWindows | Where-Object { $script:processIds -contains $_.Current.ProcessId })
}

function Find-UploadAction([string]$NamePattern) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $found = @(foreach ($root in @(Get-CopilotRoots)) {
                foreach ($control in @(Get-Descendants $root)) {
                    if ($control.Current.IsOffscreen -or -not $control.Current.IsEnabled) { continue }
                    if ($control.Current.ControlType -notin @([System.Windows.Automation.ControlType]::Button, [System.Windows.Automation.ControlType]::MenuItem)) { continue }
                    if ($control.Current.Name -match $NamePattern) { $control }
                }
            })
        if ($found.Count -eq 1) { return $found[0] }
        if ($found.Count -gt 1) {
            $final = @($found | Where-Object { $_.Current.Name -match $UploadMenuPattern })
            if ($final.Count -eq 1) { return $final[0] }
            throw "More than one Copilot control matches '$NamePattern'."
        }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMilliseconds)
    throw "Copilot control not found: $NamePattern"
}

function Invoke-UploadAction($Control) {
    $pattern = $null
    if ($Control.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
        # Invoke blocks while a modal file dialog is open. Run it in a separate
        # runspace so this thread can fill in the dialog.
        $worker = [PowerShell]::Create()
        [void]$worker.AddScript('param($p) $p.Invoke()').AddArgument($pattern)
        $script:invocations.Add([pscustomobject]@{ Worker = $worker; Async = $worker.BeginInvoke() })
        return
    }
    if ($Control.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) { $pattern.Expand(); return }
    throw 'Upload control supports neither Invoke nor Expand.'
}

function Get-FileDialogs {
    @(foreach ($window in @(Get-TopLevelWindows)) {
            if ($window.Current.ClassName -eq '#32770') { $window }
            if ($script:processIds -contains $window.Current.ProcessId) {
                Get-Descendants $window | Where-Object { $_.Current.ClassName -eq '#32770' }
            }
        }) | Sort-Object -Property @{ Expression = { $_.Current.NativeWindowHandle } } -Unique
}

function Find-FileDialog([int[]]$ExistingHandles) {
    # A new #32770 dialog owned by the Copilot window, with exactly one file name field.
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $candidates = @(foreach ($window in @(Get-FileDialogs)) {
                $handle = $window.Current.NativeWindowHandle
                if ($ExistingHandles -contains $handle) { continue }
                $owned = $script:processIds -contains $window.Current.ProcessId
                $owner = [IntPtr]$handle
                for ($depth = 0; $depth -lt 12 -and $owner -ne [IntPtr]::Zero; $depth++) {
                    if ($owner -eq $script:windowHandle) { $owned = $true; break }
                    $owner = [ComplaintDossierUploadWin32]::GetWindow($owner, $GW_OWNER)
                }
                if (-not $owned) { continue }
                $edits = @(Get-Descendants $window | Where-Object {
                        if (($_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Edit -or $_.Current.ClassName -eq 'Edit') -and -not $_.Current.IsOffscreen) {
                            $parent = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($_)
                            $_.Current.AutomationId -eq '1148' -or $_.Current.Name -match '(?i)^(file name|filename|dateiname)\s*:?' -or
                            ($null -ne $parent -and $parent.Current.AutomationId -eq '1148')
                        } else { $false }
                    })
                if ($edits.Count -eq 1) { [pscustomobject]@{ Root = $window; Edit = $edits[0] } }
            })
        if ($candidates.Count -eq 1) { return $candidates[0] }
        if ($candidates.Count -gt 1) { throw 'More than one new file dialog belongs to Copilot.' }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMilliseconds)
    throw 'No file dialog of the Copilot window appeared.'
}

function Set-DialogFileName($Dialog, [string]$Path) {
    $value = $null
    if ($Dialog.Edit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$value)) { $value.SetValue($Path); return }
    # Some dialogs expose only legacy panes. Address the verified Edit by HWND;
    # never type into whatever window is active.
    $handle = [IntPtr]$Dialog.Edit.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero -or $Dialog.Edit.Current.ClassName -ne 'Edit') { throw 'File name field cannot be controlled.' }
    $result = [UIntPtr]::Zero
    if ([ComplaintDossierUploadWin32]::SendMessageTimeout($handle, $WM_SETTEXT, [UIntPtr]::Zero, $Path, 2, 5000, [ref]$result) -eq [IntPtr]::Zero -or $result -eq [UIntPtr]::Zero) {
        throw 'Could not write the file name into the dialog.'
    }
}

function Confirm-FileDialog($Dialog) {
    $open = @(Get-Descendants $Dialog.Root | Where-Object {
            $_.Current.AutomationId -eq '1' -and ($_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Button -or $_.Current.ClassName -eq 'Button')
        })
    if ($open.Count -ne 1) { throw 'Open button in the file dialog is not unique.' }
    if ($open[0].Current.ControlType -eq [System.Windows.Automation.ControlType]::Button) { Invoke-UploadAction $open[0]; return }
    [void][ComplaintDossierUploadWin32]::SetForegroundWindow([IntPtr]$Dialog.Root.Current.NativeWindowHandle)
    $result = [UIntPtr]::Zero
    if ([ComplaintDossierUploadWin32]::SendMessageTimeout([IntPtr]$open[0].Current.NativeWindowHandle, $BM_CLICK, [UIntPtr]::Zero, [IntPtr]::Zero, 2, 5000, [ref]$result) -eq [IntPtr]::Zero) {
        throw 'Open button in the file dialog does not respond.'
    }
}

function Wait-AttachmentConfirmed([string]$Name) {
    # Confirmed = a visible control names the file together with "remove",
    # nothing is uploading, three polls in a row.
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $stable = 0
    do {
        $confirmed = $false
        $busy = $false
        foreach ($root in @(Get-CopilotRoots)) {
            foreach ($control in @(Get-Descendants $root)) {
                if ($control.Current.IsOffscreen) { continue }
                $label = [string]$control.Current.Name
                if ($label -match '(?i)upload failed|could not upload|hochladen fehlgeschlagen|konnte nicht hochgeladen|nicht unterst.tzt') { throw "Copilot rejected the upload: $Name" }
                if ($control.Current.ControlType -eq [System.Windows.Automation.ControlType]::ProgressBar -or
                    $label -match '(?i)uploading|wird hochgeladen|hochladen\.\.\.|processing|wird verarbeitet') { $busy = $true }
                if ($label.IndexOf($Name, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and $label -match '(?i)remove|entfernen') { $confirmed = $true }
            }
        }
        if ($confirmed -and -not $busy) { $stable++ } else { $stable = 0 }
        if ($stable -ge 3) { return }
        Start-Sleep -Milliseconds 200
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMilliseconds)
    throw "Upload not confirmed as a chat attachment: $Name"
}

$files = @($FilePath | ForEach-Object {
        $file = Get-Item -LiteralPath $_
        if ($file -isnot [IO.FileInfo] -or $file.Length -le 0) { throw "Upload needs an existing, non-empty file: $_" }
        $file
    })
if (@($files | Group-Object Name | Where-Object { $_.Count -gt 1 }).Count) { throw 'Two upload files share a name; the chat could not tell them apart.' }
$windows = @(Get-Process | Where-Object { $_.MainWindowTitle -eq $WindowTitle -and $_.MainWindowHandle -ne 0 })
if ($windows.Count -ne 1) { throw "Expected exactly one window titled '$WindowTitle', found $($windows.Count)." }
$script:windowHandle = $windows[0].MainWindowHandle
$script:processIds = @([int]$windows[0].Id)
$all = @(Get-CimInstance Win32_Process)
do {
    $before = $script:processIds.Count
    $script:processIds = @($script:processIds + @($all | Where-Object { $script:processIds -contains [int]$_.ParentProcessId } | ForEach-Object { [int]$_.ProcessId }) | Sort-Object -Unique)
} while ($before -ne $script:processIds.Count)
$script:invocations = New-Object 'System.Collections.Generic.List[object]'

try {
    foreach ($file in $files) {
        [void][ComplaintDossierUploadWin32]::SetForegroundWindow($script:windowHandle)
        Invoke-UploadAction (Find-UploadAction '(?i)^(quellen hinzuf.gen und verwalten|quellen|add sources|manage sources|add content|inhalt hinzuf.gen)$')
        $menu = Find-UploadAction '(?i)^(hochladen|upload|bilder und dateien hochladen|upload images and files|upload from device)$'
        if ($menu.Current.Name -match '(?i)^(hochladen|upload)$') {
            Invoke-UploadAction $menu
            $menu = Find-UploadAction $UploadMenuPattern
        }
        $existing = @(Get-TopLevelWindows | ForEach-Object { $_.Current.NativeWindowHandle }) + @(Get-FileDialogs | ForEach-Object { $_.Current.NativeWindowHandle })
        Invoke-UploadAction $menu
        $dialog = Find-FileDialog -ExistingHandles $existing
        Set-DialogFileName $dialog $file.FullName
        Confirm-FileDialog $dialog
        Wait-AttachmentConfirmed -Name $file.Name
        Write-Host ("[upload] confirmed: {0}" -f $file.Name)
    }
} finally {
    foreach ($invocation in $script:invocations) {
        if ($invocation.Async.IsCompleted) {
            try { [void]$invocation.Worker.EndInvoke($invocation.Async) } finally { $invocation.Worker.Dispose() }
        } else {
            # Do not block on an unresponsive UI provider.
            [void]$invocation.Worker.BeginStop($null, $null)
        }
    }
}
