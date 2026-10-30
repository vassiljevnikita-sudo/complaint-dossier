<#
.SYNOPSIS
  One prompt round trip through the Microsoft 365 Copilot desktop app.

.DESCRIPTION
  Started by Invoke-CopilotHandoff (src/ComplaintDossier.psm1) in Windows
  PowerShell 5.1 with -STA. Steps:
    1. check that the desktop is unlocked and find the Copilot window
    2. open a new chat, so no earlier conversation leaks into the answer;
       with -TemporaryChat switch it to a temporary chat that Copilot does
       not keep in the chat history
    3. select the configured model; abort if the menu does not offer it
    4. attach the files listed in -AttachmentListFile (Add-CopilotAttachment.ps1)
    5. type the prompt into the chat input and invoke the send button
    6. wait until the answer is complete, read its text through UI Automation
       and write it to -OutputFile (the copy button is only a fallback)
  Nothing needs the foreground window or a drawn desktop: the steps use UI
  Automation patterns and window messages to the web view. Bringing the
  window to the front is attempted, but not required.
  Every step is recorded in -StatusFile (JSON), also when the run fails.
  Exit code 0 = answer written, 1 = aborted. ASCII only (Windows PowerShell 5.1).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PromptFile,
    [Parameter(Mandatory = $true)][string]$OutputFile,
    [Parameter(Mandatory = $true)][string]$StatusFile,
    [string]$AttachmentListFile = '',
    [string]$WindowTitle = 'Microsoft 365 Copilot',
    [string]$PreferredModel = '',
    [string]$PreferredModelGroup = '',
    [switch]$TemporaryChat,
    [int]$StartDelaySeconds = 3,
    [int]$NewChatTimeoutSeconds = 30,
    [int]$UploadTimeoutSeconds = 60,
    [int]$ResponseTimeoutSeconds = 240,
    [int]$StableReadyChecks = 3,
    [int]$PollMilliseconds = 1000,
    [int]$ModelMenuTimeoutMilliseconds = 3000
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'CopilotAutomation.ps1')

$utf8 = New-Object System.Text.UTF8Encoding($false)
$status = [ordered]@{
    startedAtUtc    = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    finishedAtUtc   = $null
    windowTitle     = $WindowTitle
    preferredModel  = $PreferredModel
    modelGroup      = $PreferredModelGroup
    modelShown      = $null
    temporaryChat   = [bool]$TemporaryChat
    uploads         = @()
    inputDesktop    = $null
    foreground      = $null
    inputMethod     = $null
    answerSource    = $null
    responseChars   = 0
    steps           = @()
    error           = $null
}
function Add-Step([string]$Name) {
    $status.steps += [ordered]@{ atUtc = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'); step = $Name }
    Write-Host "[copilot] $Name"
}

$exitCode = 1
try {
    if ([string]::IsNullOrWhiteSpace($PreferredModel)) { throw 'No model configured (Copilot.PreferredModel). Refusing to run with an unknown model.' }
    $prompt = [IO.File]::ReadAllText($PromptFile, [Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($prompt)) { throw "Prompt file is empty: $PromptFile" }
    $uploads = @()
    if ($AttachmentListFile) {
        $uploads = @([IO.File]::ReadAllLines($AttachmentListFile, [Text.Encoding]::UTF8) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    Initialize-CopilotAutomation
    for ($i = $StartDelaySeconds; $i -ge 1; $i--) { Write-Host "[copilot] start in $i s"; Start-Sleep -Seconds 1 }
    $status.inputDesktop = Test-InputDesktop
    $copilot = Get-CopilotWindow $WindowTitle
    $script:CopilotProcessIds = @(Get-CopilotProcessIds ([int]$copilot.Id))
    $status.foreground = Activate-CopilotWindow $copilot
    Add-Step "window found: $($copilot.MainWindowTitle) (input desktop: $($status.inputDesktop), foreground: $($status.foreground))"

    $newChat = Find-NewChatButton $copilot.MainWindowHandle
    if ($null -eq $newChat) { throw 'New chat button not found. Aborting, so the prompt does not land in an existing conversation.' }
    Click-UiElement $newChat
    Start-Sleep -Milliseconds 800
    Wait-CopilotReady $copilot $NewChatTimeoutSeconds $StableReadyChecks $PollMilliseconds
    Add-Step 'new chat opened'
    if ($TemporaryChat) {
        Enable-TemporaryChat $copilot
        Wait-CopilotReady $copilot $NewChatTimeoutSeconds $StableReadyChecks $PollMilliseconds
        Add-Step 'temporary chat switched on'
    }

    $status.modelShown = Select-PreferredModel -Process $copilot -Target $PreferredModel -Group $PreferredModelGroup -TimeoutMilliseconds $ModelMenuTimeoutMilliseconds
    Wait-CopilotReady $copilot $NewChatTimeoutSeconds $StableReadyChecks $PollMilliseconds
    Add-Step "model selected: $PreferredModel"

    if ($uploads.Count) {
        & (Join-Path $PSScriptRoot 'Add-CopilotAttachment.ps1') -FilePath $uploads -WindowTitle $copilot.MainWindowTitle -TimeoutMilliseconds ($UploadTimeoutSeconds * 1000)
        $status.uploads = @($uploads | ForEach-Object { [IO.Path]::GetFileName($_) })
        Wait-CopilotReady $copilot $NewChatTimeoutSeconds $StableReadyChecks $PollMilliseconds
        Add-Step "$($uploads.Count) file(s) attached"
    }

    $baseline = @(Get-AnswerTitles $copilot.MainWindowHandle).Count
    $status.inputMethod = Send-CopilotPrompt $copilot $prompt
    Add-Step "prompt sent ($($status.inputMethod))"
    Start-Sleep -Seconds 3

    $answer = Wait-CopilotResponse $copilot $baseline $ResponseTimeoutSeconds $StableReadyChecks $PollMilliseconds
    $response = $answer.Text
    $status.answerSource = 'UIAutomation'
    if ([string]::IsNullOrWhiteSpace($response)) {
        if ($null -eq $answer.CopyButton) { throw 'The answer has no readable text and no copy button.' }
        $response = Copy-CopilotResponse $answer.CopyButton $prompt
        $status.answerSource = 'CopyButton'
    }
    if (Test-PromptEcho $response $prompt) { throw 'The text read from the chat is the prompt, not the answer.' }
    [IO.File]::WriteAllText($OutputFile, $response, $utf8)
    $status.responseChars = $response.Length
    Add-Step "response read ($($status.answerSource))"
    $exitCode = 0
} catch {
    $status.error = $_.Exception.Message
    Write-Host "[copilot] ABORTED: $($status.error)"
} finally {
    $status.finishedAtUtc = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    [IO.File]::WriteAllText($StatusFile, (($status | ConvertTo-Json -Depth 5) + "`n"), $utf8)
}
exit $exitCode
