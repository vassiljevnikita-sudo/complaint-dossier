# UI Automation helpers for the Microsoft 365 Copilot desktop app.
# Dot-sourced by Invoke-CopilotRoundtrip.ps1 and by the Pester tests.
# Runs in Windows PowerShell 5.1 (.NET Framework UI Automation). ASCII only.
# Defining the functions needs no Windows types; call Initialize-CopilotAutomation
# before using them against a real window.

function Initialize-CopilotAutomation {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    if (-not ('ComplaintDossierWin32' -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ComplaintDossierWin32 {
    [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
    [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
}
"@
    }
}

function Assert-InputDesktop {
    # UI Automation needs an unlocked, connected desktop session.
    $desktop = [ComplaintDossierWin32]::OpenInputDesktop(0, $false, 1)
    if ($desktop -eq [IntPtr]::Zero) { throw 'Windows input desktop not reachable. Unlock or reconnect the session; Copilot UI Automation needs an active desktop.' }
    [void][ComplaintDossierWin32]::CloseDesktop($desktop)
}

function Get-CopilotWindow {
    param([string]$Title)
    $windows = @(Get-Process | Where-Object { $_.MainWindowHandle -ne 0 })
    $exact = @($windows | Where-Object { $_.MainWindowTitle -eq $Title })
    if ($exact.Count -gt 0) { return $exact | Sort-Object StartTime -Descending | Select-Object -First 1 }
    $escaped = [regex]::Escape($Title)
    $prefixed = @($windows | Where-Object { $_.MainWindowTitle -match ('(?i)^' + $escaped + '(?:\s*[-|].*)?$') })
    if ($prefixed.Count -gt 0) { return $prefixed | Sort-Object StartTime -Descending | Select-Object -First 1 }
    $visible = @($windows | ForEach-Object { $_.MainWindowTitle } | Where-Object { $_ } | Sort-Object -Unique)
    throw ("Copilot window '{0}' not found. Open the Microsoft 365 Copilot app and try again. Visible windows: {1}" -f $Title, ($visible -join ' | '))
}

function Get-CopilotProcessIds {
    # Copilot renders in child processes (WebView2); UI elements may belong to them.
    param([int]$RootProcessId)
    $ids = New-Object 'System.Collections.Generic.HashSet[int]'
    [void]$ids.Add($RootProcessId)
    $rows = @()
    try { $rows = @(Get-CimInstance Win32_Process -ErrorAction Stop) } catch { return @($RootProcessId) }
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($row in $rows) {
            if ($ids.Contains([int]$row.ParentProcessId) -and $ids.Add([int]$row.ProcessId)) { $changed = $true }
        }
    }
    @($ids | Sort-Object)
}

function Activate-CopilotWindow {
    param([System.Diagnostics.Process]$Process)
    if ($Process.MainWindowHandle -eq 0) { throw 'Copilot has no window handle.' }
    [void][ComplaintDossierWin32]::ShowWindowAsync($Process.MainWindowHandle, 9)
    Start-Sleep -Milliseconds 250
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        [void][ComplaintDossierWin32]::SetForegroundWindow($Process.MainWindowHandle)
        Start-Sleep -Milliseconds 250
        if ([ComplaintDossierWin32]::GetForegroundWindow() -eq $Process.MainWindowHandle) { return }
        # Windows only lets a background process take the foreground after a
        # key event; a bare ALT press satisfies that rule without side effects.
        [ComplaintDossierWin32]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
        [ComplaintDossierWin32]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
        [void][ComplaintDossierWin32]::SetForegroundWindow($Process.MainWindowHandle)
        Start-Sleep -Milliseconds 250
        if ([ComplaintDossierWin32]::GetForegroundWindow() -eq $Process.MainWindowHandle) { return }
        try { [void](New-Object -ComObject WScript.Shell).AppActivate($Process.Id); Start-Sleep -Milliseconds 300 } catch {}
        if ([ComplaintDossierWin32]::GetForegroundWindow() -eq $Process.MainWindowHandle) { return }
    }
    throw 'Could not bring the Copilot window to the foreground.'
}

function Normalize-UiName([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    ($Text.ToLowerInvariant()) -replace '[^a-z0-9.]', ''
}

function Test-UiNameExact([string]$Name, [string]$Target) { (Normalize-UiName $Name) -eq (Normalize-UiName $Target) }

function Get-WindowRoot {
    param([System.Diagnostics.Process]$Process)
    try { [System.Windows.Automation.AutomationElement]::FromHandle($Process.MainWindowHandle) } catch { $null }
}

function Find-UiElements {
    # All visible, enabled descendants of the Copilot window, optionally of one control type.
    param([IntPtr]$Handle, $ControlType)
    try {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle($Handle)
        if ($null -eq $root) { return @() }
        $condition = [System.Windows.Automation.Condition]::TrueCondition
        if ($null -ne $ControlType) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ControlType)
        }
        @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition))
    } catch { @() }
}

function Find-ModelSwitcher {
    param([System.Diagnostics.Process]$Process)
    $root = Get-WindowRoot $Process
    if ($null -eq $root) { return $null }
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, 'gptModeSwitcher')
    foreach ($el in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)) {
        if (-not $el.Current.IsOffscreen -and $el.Current.IsEnabled) { return [pscustomobject]@{ Element = $el; Name = [string]$el.Current.Name } }
    }
    $null
}

function Find-NamedUiElement {
    # Exact name match inside the Copilot process tree. Buttons and menu items
    # win over plain text labels with the same name.
    param([System.Diagnostics.Process]$Process, [string]$Target)
    $root = Get-WindowRoot $Process
    if ($null -eq $root) { return $null }
    $process = @($script:CopilotProcessIds | ForEach-Object {
            New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$_)
        })
    $processCondition = $process[0]
    if ($process.Count -gt 1) { $processCondition = [System.Windows.Automation.OrCondition]::new([System.Windows.Automation.Condition[]]$process) }
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Target)
    $condition = New-Object System.Windows.Automation.AndCondition($nameCondition, $processCondition)
    $best = $null
    foreach ($el in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)) {
        try {
            if ($el.Current.IsOffscreen -or -not $el.Current.IsEnabled) { continue }
            $rect = $el.Current.BoundingRectangle
            if ($rect.Width -le 0 -or $rect.Height -le 0) { continue }
            $score = [int]$rect.Bottom
            if ($el.Current.ControlType -ne [System.Windows.Automation.ControlType]::Text) { $score += 100000 }
            if ($null -eq $best -or $score -gt $best.Score) { $best = [pscustomobject]@{ Element = $el; Name = ([string]$el.Current.Name).Trim(); Score = $score } }
        } catch {}
    }
    $best
}

function Wait-ForNamedUiElement {
    param([System.Diagnostics.Process]$Process, [string]$Target, [int]$TimeoutMilliseconds)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.ElapsedMilliseconds -lt $TimeoutMilliseconds) {
        $hit = Find-NamedUiElement $Process $Target
        if ($null -ne $hit) { return $hit }
        Start-Sleep -Milliseconds 100
    }
    $null
}

function Move-MouseToUiElement {
    param($Element)
    $rect = $Element.Current.BoundingRectangle
    if ($Element.Current.IsOffscreen -or $rect.Width -le 0 -or $rect.Height -le 0) { throw 'UI element is not visible.' }
    [void][ComplaintDossierWin32]::SetCursorPos([int]($rect.Left + $rect.Width / 2), [int]($rect.Top + $rect.Height / 2))
}

function Click-UiElement {
    # Prefer UI Automation patterns; fall back to a mouse click on the element centre.
    param($Element)
    if ($null -eq $Element) { throw 'UI element is null.' }
    $pattern = $null
    try {
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) { $pattern.Invoke(); return }
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) { $pattern.Select(); return }
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) { $pattern.Expand(); return }
    } catch {}
    Move-MouseToUiElement $Element
    Start-Sleep -Milliseconds 100
    [ComplaintDossierWin32]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    [ComplaintDossierWin32]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
}

function Select-PreferredModel {
    <#
      Selects the configured model in the Copilot model menu. If the model (or
      its group submenu) is not offered, it throws. It never falls back to
      another model.
    #>
    param([System.Diagnostics.Process]$Process, [string]$Target, [string]$Group, [int]$TimeoutMilliseconds = 3000)
    if ([string]::IsNullOrWhiteSpace($Target)) { throw 'No model configured (Copilot.PreferredModel). Refusing to run with an unknown model.' }
    Activate-CopilotWindow $Process
    $switcher = Find-ModelSwitcher $Process
    if ($null -eq $switcher) { throw 'Copilot model switcher not found.' }
    if (Test-UiNameExact $switcher.Name $Target) { return $switcher.Name }
    Click-UiElement $switcher.Element

    $selected = Wait-ForNamedUiElement $Process $Target $TimeoutMilliseconds
    if ($null -eq $selected -and -not [string]::IsNullOrWhiteSpace($Group)) {
        $parent = Wait-ForNamedUiElement $Process $Group $TimeoutMilliseconds
        if ($null -eq $parent) { throw "Configured model group not found in the Copilot menu: $Group" }
        # The group opens its submenu on hover, as in the Copilot app.
        Move-MouseToUiElement $parent.Element
        $selected = Wait-ForNamedUiElement $Process $Target $TimeoutMilliseconds
    }
    if ($null -eq $selected) { throw "Configured model not found in the Copilot menu: $Target" }
    Click-UiElement $selected.Element
    Start-Sleep -Milliseconds 700
    Activate-CopilotWindow $Process
    $after = Find-ModelSwitcher $Process
    $shown = ''
    if ($null -ne $after) { $shown = $after.Name }
    Write-Host ("[model] selected '{0}' (switcher now shows '{1}')" -f $Target, $shown)
    $shown
}

function Find-CopilotInput {
    # The chat input is the widest, lowest edit control in the window.
    param([IntPtr]$Handle)
    $best = $null
    foreach ($el in @(Find-UiElements $Handle ([System.Windows.Automation.ControlType]::Edit))) {
        try {
            if (-not $el.Current.IsEnabled -or $el.Current.IsOffscreen) { continue }
            $r = $el.Current.BoundingRectangle
            if ($r.Width -lt 200 -or $r.Height -lt 20) { continue }
            $score = [int]$r.Bottom + [int]($r.Width / 10)
            if ([string]$el.Current.Name -match '(?i)message|nachricht|ask|fragen|prompt|copilot|chat') { $score += 1000 }
            if ($null -eq $best -or $score -gt $best.Score) { $best = [pscustomobject]@{ Element = $el; Score = $score } }
        } catch {}
    }
    if ($null -eq $best) { return $null }
    $best.Element
}

function Get-InputValue {
    param($InputElement)
    $pattern = $null
    try {
        if ($InputElement.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) { return [string]$pattern.Current.Value }
        if ($InputElement.TryGetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern, [ref]$pattern)) { return [string]$pattern.DocumentRange.GetText(-1) }
    } catch {}
    $null
}

function Test-CopilotBusy {
    param([IntPtr]$Handle)
    foreach ($button in @(Find-UiElements $Handle ([System.Windows.Automation.ControlType]::Button))) {
        try { if ([string]$button.Current.Name -match '(?i)stop generating|stop responding|antwort beenden|generierung beenden|antwort abbrechen') { return $true } } catch {}
    }
    $false
}

function Get-CopilotUiState {
    # Ready = input enabled and empty, no "stop generating" button.
    param([IntPtr]$Handle)
    $busy = Test-CopilotBusy $Handle
    $box = Find-CopilotInput $Handle
    if ($null -eq $box) { return [pscustomobject]@{ Busy = $busy; Ready = $false; Text = 'input not found' } }
    $value = Get-InputValue $box
    $empty = ($null -eq $value) -or [string]::IsNullOrWhiteSpace($value)
    $ready = $box.Current.IsEnabled -and -not $busy -and $empty
    [pscustomobject]@{ Busy = $busy; Ready = $ready; Text = ('busy={0} inputEmpty={1} ready={2}' -f $busy, $empty, $ready) }
}

function Wait-CopilotReady {
    param([System.Diagnostics.Process]$Process, [int]$TimeoutSeconds, [int]$StableChecks, [int]$PollMilliseconds)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $stable = 0
    $state = [pscustomobject]@{ Text = 'not checked' }
    while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $state = Get-CopilotUiState $Process.MainWindowHandle
        if ($state.Ready) { $stable++; if ($stable -ge $StableChecks) { return } } else { $stable = 0 }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
    throw "Copilot did not become ready within $TimeoutSeconds s ($($state.Text))."
}

function Find-NewChatButton {
    param([IntPtr]$Handle)
    foreach ($type in @([System.Windows.Automation.ControlType]::Button, [System.Windows.Automation.ControlType]::MenuItem, [System.Windows.Automation.ControlType]::Hyperlink)) {
        foreach ($el in @(Find-UiElements $Handle $type)) {
            try {
                if ($el.Current.IsEnabled -and -not $el.Current.IsOffscreen -and ([string]$el.Current.Name).Trim() -match '(?i)^(new chat|new conversation|neuer chat|neue unterhaltung)(\b.*)?$') { return $el }
            } catch {}
        }
    }
    $null
}

function Get-CopyButtons {
    param([IntPtr]$Handle)
    @(foreach ($button in @(Find-UiElements $Handle ([System.Windows.Automation.ControlType]::Button))) {
            try {
                if ($button.Current.IsOffscreen -or -not $button.Current.IsEnabled) { continue }
                if (([string]$button.Current.Name).Trim() -notmatch '(?i)^(copy|copy response|kopieren|antwort kopieren)$') { continue }
                [pscustomobject]@{ Element = $button; Bottom = [double]$button.Current.BoundingRectangle.Bottom }
            } catch {}
        })
}

function Get-TextHash([string]$Text) {
    $normalized = ''
    if ($Text) { $normalized = $Text.Replace("`r`n", "`n").Trim() }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($normalized)))).Replace('-', '') } finally { $sha.Dispose() }
}

function Send-CopilotPrompt {
    <#
      Puts the prompt into the chat input and presses Enter. Tries ValuePattern
      first and checks the read-back; falls back to clipboard paste. The prompt
      is never left in the clipboard.
    #>
    param([System.Diagnostics.Process]$Process, [string]$Text)
    Activate-CopilotWindow $Process
    $box = Find-CopilotInput $Process.MainWindowHandle
    if ($null -eq $box) { throw 'Copilot chat input not found.' }
    $box.SetFocus()
    Start-Sleep -Milliseconds 120
    $expected = Get-TextHash $Text
    $method = $null
    $pattern = $null
    try {
        if ($box.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
            $pattern.SetValue($Text)
            Start-Sleep -Milliseconds 300
            if ((Get-TextHash (Get-InputValue $box)) -eq $expected) { $method = 'ValuePattern' }
        }
    } catch {}
    if ($null -eq $method) {
        $box.SetFocus()
        [System.Windows.Forms.SendKeys]::SendWait('^a')
        Set-Clipboard -Value $Text
        Start-Sleep -Milliseconds 80
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds 300
        Set-Clipboard -Value ' '
        $method = 'ClipboardPaste'
    }
    Start-Sleep -Milliseconds 180
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Milliseconds 200
    $method
}

function Wait-CopilotResponse {
    # Done when a new copy button exists and the input is ready for several polls.
    param([System.Diagnostics.Process]$Process, [int]$BaselineCopyButtons, [int]$TimeoutSeconds, [int]$StableChecks, [int]$PollMilliseconds)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $stable = 0
    while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $Process.Refresh()
        if ($Process.HasExited -or $Process.MainWindowHandle -eq 0) { throw 'Copilot window closed while waiting for the answer.' }
        $state = Get-CopilotUiState $Process.MainWindowHandle
        $buttons = @(Get-CopyButtons $Process.MainWindowHandle)
        $newCopy = $null
        if ($buttons.Count -gt $BaselineCopyButtons) { $newCopy = $buttons | Sort-Object Bottom -Descending | Select-Object -First 1 }
        if ($state.Ready -and $null -ne $newCopy) {
            $stable++
            if ($stable -ge $StableChecks) { return $newCopy }
        } else { $stable = 0 }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
    throw "No finished Copilot answer within $TimeoutSeconds s."
}

function Copy-CopilotResponse {
    <#
      Clicks the copy button of the newest answer and reads the clipboard.
      A sentinel proves that the click changed the clipboard; a copy of the
      prompt is rejected.
    #>
    param($CopyButton, [string]$Prompt)
    $sentinel = '__COMPLAINT_DOSSIER_WAITING_' + [guid]::NewGuid().ToString('N') + '__'
    Set-Clipboard -Value $sentinel
    Start-Sleep -Milliseconds 80
    Click-UiElement $CopyButton.Element
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.ElapsedMilliseconds -lt 5000) {
        $text = $null
        try { $text = [string](Get-Clipboard -Raw) } catch {}
        if (-not [string]::IsNullOrWhiteSpace($text) -and $text -ne $sentinel) {
            if ((Get-TextHash $text) -eq (Get-TextHash $Prompt) -or $text.Contains('### Statistics (calculated by the tool)')) {
                throw 'The copied text is the prompt, not the answer.'
            }
            return $text
        }
        Start-Sleep -Milliseconds 100
    }
    throw 'Copy button clicked, but no answer arrived in the clipboard.'
}
