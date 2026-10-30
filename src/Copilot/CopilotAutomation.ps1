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
    [DllImport("user32.dll", EntryPoint="SendMessageTimeoutW")] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
    public static bool Send(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam) { IntPtr r; return SendMessageTimeout(hWnd, msg, wParam, lParam, 2, 5000, out r) != IntPtr.Zero; }
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);
    [DllImport("user32.dll")] public static extern bool GetKeyboardState(byte[] state);
    [DllImport("user32.dll")] public static extern bool SetKeyboardState(byte[] state);
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder name, int count);
    public static System.Collections.Generic.List<IntPtr> FindChildren(IntPtr parent, string className) {
        var found = new System.Collections.Generic.List<IntPtr>();
        EnumChildWindows(parent, (h, l) => { var sb = new System.Text.StringBuilder(256); GetClassName(h, sb, 256); if (sb.ToString() == className) found.Add(h); return true; }, IntPtr.Zero);
        return found;
    }
}
"@
    }
}

function Test-InputDesktop {
    # Reports whether the session has a reachable input desktop (connected and
    # unlocked). Informational only: the round trip uses UI Automation and
    # window messages addressed to the Copilot window, never input to whatever
    # window is active, so it also runs in a disconnected or undrawn session.
    $desktop = [ComplaintDossierWin32]::OpenInputDesktop(0, $false, 1)
    if ($desktop -eq [IntPtr]::Zero) { return $false }
    [void][ComplaintDossierWin32]::CloseDesktop($desktop)
    $true
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
    # Opportunistic: tries to bring the window to the front and reports whether
    # that worked. Nothing depends on it. Every later step uses UI Automation
    # patterns or window messages, which also work when the window is not in the
    # foreground or the remote desktop session is not drawn.
    param([System.Diagnostics.Process]$Process)
    if ($Process.MainWindowHandle -eq 0) { throw 'Copilot has no window handle.' }
    try {
        [void][ComplaintDossierWin32]::ShowWindowAsync($Process.MainWindowHandle, 9)
        [void][ComplaintDossierWin32]::SetForegroundWindow($Process.MainWindowHandle)
        Start-Sleep -Milliseconds 150
        return ([ComplaintDossierWin32]::GetForegroundWindow() -eq $Process.MainWindowHandle)
    } catch { return $false }
}

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


function Move-MouseToUiElement {
    param($Element)
    $rect = $Element.Current.BoundingRectangle
    if ($Element.Current.IsOffscreen -or $rect.Width -le 0 -or $rect.Height -le 0) { throw 'UI element is not visible.' }
    [void][ComplaintDossierWin32]::SetCursorPos([int]($rect.Left + $rect.Width / 2), [int]($rect.Top + $rect.Height / 2))
}


function Click-UiElement {
    # UI Automation patterns first (they need no foreground); a mouse click on the
    # element centre is the last resort and only works on a drawn desktop.
    param($Element)
    if ($null -eq $Element) { throw 'UI element is null.' }
    $pattern = $null
    try {
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) { $pattern.Invoke(); return }
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$pattern)) { $pattern.Toggle(); return }
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) { $pattern.Select(); return }
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) { $pattern.Expand(); return }
    } catch {}
    Move-MouseToUiElement $Element
    Start-Sleep -Milliseconds 100
    [ComplaintDossierWin32]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    [ComplaintDossierWin32]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
}

function Test-ModelLabel {
    # Model menu items read "<model> <description>" ("Think Deeper Think longer
    # for better answers"), the switcher reads "<model>, <hint>". A label
    # matches when it is the model name or starts with it plus a space or comma.
    param([string]$Name, [string]$Target)
    $n = ([string]$Name).Trim()
    $t = ([string]$Target).Trim()
    if (-not $t) { return $false }
    $c = [StringComparison]::OrdinalIgnoreCase
    $n.Equals($t, $c) -or $n.StartsWith("$t ", $c) -or $n.StartsWith("$t,", $c)
}

function Select-ModelChoice {
    # Picks one menu entry for $Target from candidates with a Name property.
    # Exact names win over prefix matches; two equal matches are an error.
    param([object[]]$Candidates, [string]$Target)
    $exact = @($Candidates | Where-Object { ([string]$_.Name).Trim() -ieq $Target.Trim() })
    if ($exact.Count -eq 1) { return $exact[0] }
    $matching = @($Candidates | Where-Object { Test-ModelLabel $_.Name $Target })
    if ($exact.Count -gt 1 -or $matching.Count -gt 1) {
        throw "More than one Copilot menu entry matches '$Target': $(($matching | ForEach-Object { $_.Name }) -join ' | ')"
    }
    if ($matching.Count -eq 1) { return $matching[0] }
    $null
}

function Find-ModelChoice {
    # Visible menu entries (radio buttons, menu items) of the Copilot process tree.
    param([System.Diagnostics.Process]$Process, [string]$Target)
    $root = Get-WindowRoot $Process
    if ($null -eq $root) { return $null }
    $types = @([System.Windows.Automation.ControlType]::RadioButton, [System.Windows.Automation.ControlType]::MenuItem, [System.Windows.Automation.ControlType]::ListItem)
    $candidates = @(foreach ($el in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
            try {
                $c = $el.Current
                if ($c.IsOffscreen -or -not $c.IsEnabled -or $types -notcontains $c.ControlType) { continue }
                if ($script:CopilotProcessIds -notcontains $c.ProcessId) { continue }
                [pscustomobject]@{ Element = $el; Name = ([string]$c.Name).Trim() }
            } catch {}
        })
    Select-ModelChoice $candidates $Target
}

function Wait-ForModelChoice {
    param([System.Diagnostics.Process]$Process, [string]$Target, [int]$TimeoutMilliseconds)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.ElapsedMilliseconds -lt $TimeoutMilliseconds) {
        $hit = Find-ModelChoice $Process $Target
        if ($null -ne $hit) { return $hit }
        Start-Sleep -Milliseconds 100
    }
    $null
}

function Close-CopilotMenu {
    # Collapses the model switcher; without the pattern, posts Escape to the web view.
    param([System.Diagnostics.Process]$Process)
    $pattern = $null
    try {
        $switcher = Find-ModelSwitcher $Process
        if ($null -ne $switcher -and $switcher.Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) { $pattern.Collapse(); return }
    } catch {}
    try { Send-KeyMessage -Handle (Get-RenderWidget $Process.MainWindowHandle) -VirtualKey 0x1B } catch {}
}

function Open-UiSubmenu {
    # Opens a submenu entry through ExpandCollapse or Invoke; hovering with the
    # mouse is the last resort.
    param($Element)
    $pattern = $null
    try {
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) { $pattern.Expand(); return }
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) { $pattern.Invoke(); return }
    } catch {}
    Move-MouseToUiElement $Element
}

function Select-PreferredModel {
    <#
      Selects the configured model in the Copilot model menu and checks that
      the switcher shows it afterwards. If the model (or its group submenu) is
      not offered, it closes the menu and throws. It never falls back to
      another model.
    #>
    param([System.Diagnostics.Process]$Process, [string]$Target, [string]$Group, [int]$TimeoutMilliseconds = 3000)
    if ([string]::IsNullOrWhiteSpace($Target)) { throw 'No model configured (Copilot.PreferredModel). Refusing to run with an unknown model.' }
    [void](Activate-CopilotWindow $Process)
    $switcher = Find-ModelSwitcher $Process
    if ($null -eq $switcher) { throw 'Copilot model switcher not found.' }
    if (Test-ModelLabel $switcher.Name $Target) { return $switcher.Name }
    Click-UiElement $switcher.Element

    $selected = Wait-ForModelChoice $Process $Target $TimeoutMilliseconds
    if ($null -eq $selected -and -not [string]::IsNullOrWhiteSpace($Group)) {
        $parent = Wait-ForModelChoice $Process $Group $TimeoutMilliseconds
        if ($null -eq $parent) { Close-CopilotMenu $Process; throw "Configured model group not found in the Copilot menu: $Group" }
        Open-UiSubmenu $parent.Element
        $selected = Wait-ForModelChoice $Process $Target $TimeoutMilliseconds
    }
    if ($null -eq $selected) { Close-CopilotMenu $Process; throw "Configured model not found in the Copilot menu: $Target" }
    Click-UiElement $selected.Element
    Start-Sleep -Milliseconds 700
    $after = Find-ModelSwitcher $Process
    $shown = ''
    if ($null -ne $after) { $shown = $after.Name }
    if (-not (Test-ModelLabel $shown $Target)) { throw "Copilot did not confirm the model '$Target'; the switcher shows '$shown'." }
    Write-Host ("[model] selected '{0}' (switcher shows '{1}')" -f $Target, $shown)
    $shown
}

function Get-TemporaryChatToggle {
    # Returns the toggle and its state, or $null. The app re-renders the button
    # after a click, so callers look it up again instead of keeping the element.
    param([System.Diagnostics.Process]$Process)
    $buttons = @(Find-UiElements $Process.MainWindowHandle ([System.Windows.Automation.ControlType]::Button) | Where-Object {
            -not $_.Current.IsOffscreen -and ([string]$_.Current.Name).Trim() -match '(?i)^(temporary chat|tempor.rer chat)$'
        })
    if ($buttons.Count -ne 1) { return $null }
    $pattern = $null
    $state = 'unknown'
    if ($buttons[0].TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$pattern)) { $state = [string]$pattern.Current.ToggleState }
    [pscustomobject]@{ Element = $buttons[0]; State = $state }
}

function Enable-TemporaryChat {
    # Switches the new chat to a temporary chat, which Copilot does not keep in
    # the chat history. Fails if the toggle is missing or does not switch on.
    param([System.Diagnostics.Process]$Process, [int]$TimeoutMilliseconds = 5000)
    $toggle = Get-TemporaryChatToggle $Process
    if ($null -eq $toggle) { throw 'Temporary chat toggle not found. Set Copilot.TemporaryChat to false if your Copilot has none.' }
    if ($toggle.State -eq 'On') { return }
    Click-UiElement $toggle.Element
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.ElapsedMilliseconds -lt $TimeoutMilliseconds) {
        Start-Sleep -Milliseconds 250
        $toggle = Get-TemporaryChatToggle $Process
        if ($null -ne $toggle -and $toggle.State -eq 'On') { return }
    }
    $state = 'toggle not found'
    if ($null -ne $toggle) { $state = $toggle.State }
    throw "Temporary chat did not switch on (state: $state)."
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
    # Copy buttons in the chat, in document order. Offscreen buttons count too:
    # a long answer pushes its own copy button below the visible area, while the
    # copy button of the prompt above it is still on screen. Invoke works on
    # offscreen buttons, and an offscreen element may report an empty rectangle,
    # so the order of the UI Automation tree decides, not the screen position.
    param([IntPtr]$Handle)
    @(foreach ($button in @(Find-UiElements $Handle ([System.Windows.Automation.ControlType]::Button))) {
            try {
                if (-not $button.Current.IsEnabled) { continue }
                if (([string]$button.Current.Name).Trim() -notmatch '(?i)^(copy|copy response|kopieren|antwort kopieren)$') { continue }
                [pscustomobject]@{ Element = $button; Bottom = [double]$button.Current.BoundingRectangle.Bottom; Offscreen = [bool]$button.Current.IsOffscreen }
            } catch {}
        })
}
function Get-TextHash([string]$Text) {
    $normalized = ''
    if ($Text) { $normalized = $Text.Replace("`r`n", "`n").Trim() }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($normalized)))).Replace('-', '') } finally { $sha.Dispose() }
}

function Get-RenderWidget {
    # The WebView2 child window that receives keyboard messages for the chat page.
    param([IntPtr]$Handle)
    $widgets = @([ComplaintDossierWin32]::FindChildren($Handle, 'Chrome_RenderWidgetHostHWND') | Where-Object { [ComplaintDossierWin32]::IsWindowVisible($_) })
    if ($widgets.Count -eq 0) { throw 'Copilot web view window (Chrome_RenderWidgetHostHWND) not found.' }
    $widgets[0]
}

function Send-KeyMessage {
    <#
      Sends one key press to the web view (synchronous window messages, so the
      order with the typed characters is kept). Shift or Control are set in the
      keyboard state of the web view's input thread while the key is handled, so
      Shift+Enter makes a line break instead of sending. Needs no foreground
      window and no drawn desktop.
    #>
    param([IntPtr]$Handle, [int]$VirtualKey, [switch]$Shift, [switch]$Control, [switch]$WithChar)
    $processId = [uint32]0
    $target = [ComplaintDossierWin32]::GetWindowThreadProcessId($Handle, [ref]$processId)
    $me = [ComplaintDossierWin32]::GetCurrentThreadId()
    $attached = $false
    $original = $null
    if ($Shift -or $Control) {
        $attached = [ComplaintDossierWin32]::AttachThreadInput($me, $target, $true)
        $state = New-Object byte[] 256
        [void][ComplaintDossierWin32]::GetKeyboardState($state)
        $original = [byte[]]$state.Clone()
        if ($Shift) { $state[0x10] = 0x80; $state[0xA0] = 0x80 }
        if ($Control) { $state[0x11] = 0x80; $state[0xA2] = 0x80 }
        [void][ComplaintDossierWin32]::SetKeyboardState($state)
    }
    try {
        if (-not [ComplaintDossierWin32]::Send($Handle, 0x0100, [IntPtr]$VirtualKey, [IntPtr]1)) { throw 'The Copilot web view did not take a key message.' }
        if ($WithChar) { [void][ComplaintDossierWin32]::Send($Handle, 0x0102, [IntPtr]$VirtualKey, [IntPtr]1) }
        [void][ComplaintDossierWin32]::Send($Handle, 0x0101, [IntPtr]$VirtualKey, [IntPtr]0xC0000001)
        Start-Sleep -Milliseconds 30
    } finally {
        if ($null -ne $original) { [void][ComplaintDossierWin32]::SetKeyboardState($original) }
        if ($attached) { [void][ComplaintDossierWin32]::AttachThreadInput($me, $target, $false) }
    }
}

function Send-TextMessages {
    # Types text into the focused web element with WM_CHAR messages (Unicode,
    # synchronous); line breaks become Shift+Enter.
    param([IntPtr]$Handle, [string]$Text)
    $lines = $Text.Replace("`r`n", "`n").Replace("`t", '    ') -split "`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($i -gt 0) { Send-KeyMessage -Handle $Handle -VirtualKey 0x0D -Shift -WithChar }
        foreach ($ch in $lines[$i].ToCharArray()) {
            if (-not [ComplaintDossierWin32]::Send($Handle, 0x0102, [IntPtr][int]$ch, [IntPtr]1)) { throw 'The Copilot web view did not take a typed character.' }
        }
    }
}

function Get-ComparableText {
    <#
      Text as the chat input holds it, for the read-back check. The editor turns
      line breaks and Markdown marks typed at a line start (#, -, 1.) into its own
      formatting, so -Strict ignores white space only, and the default also ignores
      the Markdown mark characters. Letters, digits and punctuation must match.
    #>
    param([string]$Text, [switch]$Strict)
    if (-not $Text) { return '' }
    $flat = $Text -replace '\s+', ''
    if ($Strict) { return $flat }
    $marks = '[#*>\-' + [char]0x2022 + ']'
    ($flat -replace '\d+[.)]', '') -replace $marks, ''
}

function Clear-CopilotInput {
    # Select all and delete, repeated until the input reads empty.
    param([IntPtr]$Widget, $InputElement, [int]$TimeoutSeconds = 30)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        Send-KeyMessage -Handle $Widget -VirtualKey 0x41 -Control
        Send-KeyMessage -Handle $Widget -VirtualKey 0x08
        Start-Sleep -Milliseconds 700
        if ($null -eq $InputElement -or -not (Get-InputValue $InputElement)) { return }
    } while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds)
    throw 'Could not clear the Copilot chat input.'
}

function Find-SendButton {
    param([IntPtr]$Handle)
    @(Find-UiElements $Handle ([System.Windows.Automation.ControlType]::Button) | Where-Object {
            try { $_.Current.IsEnabled -and ([string]$_.Current.Name).Trim() -match '(?i)^(senden|send|submit|absenden|nachricht senden|send message)$' } catch { $false }
        }) | Select-Object -First 1
}

function Send-CopilotPrompt {
    <#
      Puts the prompt into the chat input and sends it, without foreground,
      clipboard or SendKeys:
        1. ValuePattern.SetValue, if the input takes it (checked by read-back)
        2. otherwise WM_CHAR messages to the web view, line breaks as Shift+Enter,
           checked by read-back of the input text
      Then the send button is invoked through UI Automation (Enter as fallback).
      Throws, without sending, when the input does not hold the prompt.
    #>
    param($Process, [string]$Text, [int]$TimeoutSeconds = 900, [int]$StallSeconds = 180)
    $box = Find-CopilotInput $Process.MainWindowHandle
    if ($null -eq $box) { throw 'Copilot chat input not found.' }
    try { $box.SetFocus() } catch {}
    Start-Sleep -Milliseconds 150
    $expected = Get-ComparableText $Text
    $method = $null
    $pattern = $null
    try {
        if ($box.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
            $pattern.SetValue($Text)
            Start-Sleep -Milliseconds 400
            if ((Get-ComparableText (Get-InputValue $box)) -eq $expected) { $method = 'ValuePattern' }
        }
    } catch {}
    $widget = Get-RenderWidget $Process.MainWindowHandle
    if ($null -eq $method) {
        if (Get-InputValue $box) { Clear-CopilotInput $widget $box }
        try { $box.SetFocus() } catch {}
        Send-TextMessages -Handle $widget -Text $Text
        # The page may take a while to process the typed text, especially when
        # Windows treats the window as hidden. Wait while the input keeps growing.
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $actual = ''
        $lastLength = -1
        $lastChange = 0
        do {
            Start-Sleep -Milliseconds 1000
            $actual = Get-ComparableText (Get-InputValue $box)
            if ($actual -eq $expected) { break }
            if ($actual.Length -ne $lastLength) { $lastLength = $actual.Length; $lastChange = $watch.Elapsed.TotalSeconds }
        } while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds -and ($watch.Elapsed.TotalSeconds - $lastChange) -lt $StallSeconds)
        if ($actual -ne $expected) {
            $at = 0
            while ($at -lt [Math]::Min($actual.Length, $expected.Length) -and $actual[$at] -eq $expected[$at]) { $at++ }
            try { Clear-CopilotInput $widget $box } catch {}
            throw ("The chat input does not hold the prompt (expected {0} characters without white space and Markdown marks, found {1}, first difference at {2}). Nothing was sent." -f $expected.Length, $actual.Length, $at)
        }
        $method = 'WindowMessages'
        if ((Get-ComparableText (Get-InputValue $box) -Strict) -ne (Get-ComparableText $Text -Strict)) { $method = 'WindowMessages (editor formatted Markdown marks)' }
    }
    $send = Find-SendButton $Process.MainWindowHandle
    if ($null -ne $send) { Click-UiElement $send; $method += '+SendButton' }
    else { Send-KeyMessage -Handle $widget -VirtualKey 0x0D -WithChar; $method += '+Enter' }
    Start-Sleep -Milliseconds 500
    $method
}

function Get-AnswerTitles {
    # "Copilot said:" labels, one per answer, in document order.
    param([IntPtr]$Handle)
    @(Find-UiElements $Handle ([System.Windows.Automation.ControlType]::Text) | Where-Object {
            try { ([string]$_.Current.AutomationId) -like 'copilot-message-*-title' } catch { $false }
        })
}

function Get-UiTextLines {
    <#
      Text of an answer, one line per paragraph or heading, from the UI
      Automation tree. Buttons (citations, file cards) and images are skipped.
      $Node is anything with Type, Name and Children; see ConvertTo-UiNode.
    #>
    param($Node, [int]$Depth = 0)
    switch ($Node.Type) {
        'Button' { return }
        'Image' { return }
        'Text' { if ([string]$Node.Name -and $Node.Name.Trim()) { return $Node.Name.Trim() }; return }
        'ListItem' {
            # List items may carry their text only as Name; an ordered list does not expose its numbers.
            $inner = (@(foreach ($child in $Node.Children) { Get-UiTextLines $child ($Depth + 1) }) -join ' ').Trim()
            if (-not $inner) { $inner = ([string]$Node.Name).Trim() }
            if ($inner) { return ('- ' + $inner) }
            return
        }
        default {
            foreach ($child in $Node.Children) { Get-UiTextLines $child ($Depth + 1) }
        }
    }
}

function ConvertTo-UiNode {
    # Plain copy of a UI Automation subtree (control view), so the text logic can be tested without Windows.
    param($Element, [int]$Depth = 0)
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $c = $Element.Current
    $children = New-Object System.Collections.Generic.List[object]
    if ($Depth -lt 30) {
        $child = $walker.GetFirstChild($Element)
        while ($null -ne $child) { $children.Add((ConvertTo-UiNode $child ($Depth + 1))); $child = $walker.GetNextSibling($child) }
    }
    [pscustomobject]@{ Type = $c.ControlType.ProgrammaticName.Replace('ControlType.', ''); Name = [string]$c.Name; Id = [string]$c.AutomationId; Children = $children.ToArray() }
}

function Get-TextNodeCount {
    param($Node)
    $n = 0
    if ($Node.Type -eq 'Text') { $n = 1 }
    foreach ($child in $Node.Children) { $n += Get-TextNodeCount $child }
    $n
}

function Get-AnswerBodyNode {
    <#
      In the answer group the first group after the "Copilot said:" label holds
      the answer body, followed by file cards and a "from <author>" line. The body
      is the child group with the most text in it; those other parts are left out.
    #>
    param($AnswerNode)
    $content = @($AnswerNode.Children | Where-Object { $_.Id -notlike 'copilot-message-*-title' -and $_.Type -eq 'Group' }) | Select-Object -First 1
    if ($null -eq $content) { return $null }
    $groups = @($content.Children | Where-Object { $_.Type -eq 'Group' })
    if ($groups.Count -eq 0) { return $content }
    $groups | Sort-Object -Property @{ Expression = { Get-TextNodeCount $_ }; Descending = $true } | Select-Object -First 1
}

function Get-CopilotAnswerText {
    # Text of the newest answer, read from UI Automation; '' when there is none.
    param([IntPtr]$Handle)
    $titles = @(Get-AnswerTitles $Handle)
    if ($titles.Count -eq 0) { return '' }
    $group = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($titles[-1])
    if ($null -eq $group) { return '' }
    $body = Get-AnswerBodyNode (ConvertTo-UiNode $group)
    if ($null -eq $body) { return '' }
    (@(Get-UiTextLines $body) -join "`n`n").Trim()
}

function Wait-CopilotResponse {
    <#
      Done when a new answer exists, Copilot is no longer generating, and the
      answer text read through UI Automation stays the same for several polls.
      Returns the copy button of the answer too, for the optional clipboard fallback.
    #>
    param([System.Diagnostics.Process]$Process, [int]$BaselineAnswers, [int]$TimeoutSeconds, [int]$StableChecks, [int]$PollMilliseconds)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $stable = 0
    $last = $null
    while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $Process.Refresh()
        if ($Process.HasExited -or $Process.MainWindowHandle -eq 0) { throw 'Copilot window closed while waiting for the answer.' }
        $handle = $Process.MainWindowHandle
        $text = ''
        if (@(Get-AnswerTitles $handle).Count -gt $BaselineAnswers -and -not (Test-CopilotBusy $handle)) { $text = Get-CopilotAnswerText $handle }
        if ($text -and $text -eq $last) {
            $stable++
            if ($stable -ge $StableChecks) {
                $buttons = @(Get-CopyButtons $handle)
                $copy = $null
                if ($buttons.Count) { $copy = $buttons[-1] }
                return [pscustomobject]@{ Text = $text; CopyButton = $copy }
            }
        } else { $stable = 0 }
        $last = $text
        Start-Sleep -Milliseconds $PollMilliseconds
    }
    throw "No finished Copilot answer within $TimeoutSeconds s."
}

function Copy-CopilotResponse {
    <#
      Optional fallback when UI Automation gives no answer text: clicks the copy
      button of the newest answer and reads the clipboard. A sentinel proves that
      the click changed the clipboard; a copy of the prompt is rejected.
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
            if (Test-PromptEcho $text $Prompt) { throw 'The copied text is the prompt, not the answer.' }
            return $text
        }
        Start-Sleep -Milliseconds 100
    }
    throw 'Copy button clicked, but no answer arrived in the clipboard.'
}

function Test-PromptEcho {
    # True when a text is the prompt itself rather than an answer to it.
    param([string]$Text, [string]$Prompt)
    (Get-TextHash $Text) -eq (Get-TextHash $Prompt) -or $Text.Contains('### Complaint record (from the complaint list)')
}
