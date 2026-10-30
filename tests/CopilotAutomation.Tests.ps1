# Pester 5 tests for the Copilot UI Automation logic.
# The model selection runs against mocks, so these tests run on any OS.
# The tests tagged 'Windows' start the real round-trip script in Windows
# PowerShell 5.1 and check that it aborts cleanly when no Copilot window exists.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    . (Join-Path $script:Root 'src/Copilot/CopilotAutomation.ps1')
}

Describe 'Model label matching' {
    It 'matches menu items with a description and the switcher with a hint' {
        Test-ModelLabel 'Think Deeper Think longer for better answers' 'Think Deeper' | Should -BeTrue
        Test-ModelLabel 'Quick Response, Modellauswahl' 'quick response' | Should -BeTrue
        Test-ModelLabel 'Auto' 'Auto' | Should -BeTrue
    }
    It 'does not match a different model with the same beginning' {
        Test-ModelLabel 'Think Deeper Think longer' 'Think Deep' | Should -BeFalse
        Test-ModelLabel 'Auto, model picker' '' | Should -BeFalse
    }
    It 'prefers an exact name and rejects two prefix matches' {
        $items = @([pscustomobject]@{ Name = 'Model A' }, [pscustomobject]@{ Name = 'Model A mini' })
        (Select-ModelChoice $items 'Model A').Name | Should -Be 'Model A'
        $items = @([pscustomobject]@{ Name = 'Model A fast' }, [pscustomobject]@{ Name = 'Model A slow' })
        { Select-ModelChoice $items 'Model A' } | Should -Throw '*More than one*'
        Select-ModelChoice $items 'Model B' | Should -BeNullOrEmpty
    }
}

Describe 'Select-PreferredModel' {
    BeforeEach {
        Mock Activate-CopilotWindow {}
        Mock Start-Sleep {}
        Mock Write-Host {}
        Mock Move-MouseToUiElement {}
        Mock Close-CopilotMenu {}
        $script:clicked = @()
        Mock Click-UiElement { $script:clicked += $Element }
        # The switcher shows the last clicked model, like the real app.
        Mock Find-ModelSwitcher {
            $name = 'Auto, model picker'
            if ($script:clicked -contains 'model-x') { $name = 'Model X, model picker' }
            [pscustomobject]@{ Element = 'switcher'; Name = $name }
        }
    }

    It 'refuses to run without a configured model' {
        { Select-PreferredModel -Process $null -Target '' } | Should -Throw '*No model configured*'
        Should -Invoke Click-UiElement -Times 0
    }
    It 'aborts when the menu does not offer the model, closes the menu and clicks nothing else' {
        Mock Wait-ForModelChoice { $null }
        { Select-PreferredModel -Process $null -Target 'Model X' } | Should -Throw '*Configured model not found in the Copilot menu: Model X*'
        $script:clicked | Should -Be @('switcher')
        Should -Invoke Close-CopilotMenu -Times 1 -Exactly
    }
    It 'aborts when the configured model group is missing' {
        Mock Wait-ForModelChoice { $null }
        { Select-PreferredModel -Process $null -Target 'Model X' -Group 'More models' } | Should -Throw '*model group not found*More models*'
    }
    It 'selects the model and returns what the switcher shows' {
        Mock Wait-ForModelChoice { if ($Target -eq 'Model X') { [pscustomobject]@{ Element = 'model-x'; Name = 'Model X Fast answers' } } }
        Select-PreferredModel -Process $null -Target 'Model X' | Should -Be 'Model X, model picker'
        $script:clicked | Should -Be @('switcher', 'model-x')
    }
    It 'opens the group submenu before looking for the model' {
        $script:hovered = $false
        Mock Move-MouseToUiElement { $script:hovered = $true }
        Mock Wait-ForModelChoice {
            if ($Target -eq 'More models') { return [pscustomobject]@{ Element = 'group'; Name = 'More models' } }
            if ($Target -eq 'Model X' -and $script:hovered) { return [pscustomobject]@{ Element = 'model-x'; Name = 'Model X' } }
            $null
        }
        Select-PreferredModel -Process $null -Target 'Model X' -Group 'More models' | Should -Be 'Model X, model picker'
        $script:hovered | Should -BeTrue
    }
    It 'fails when the switcher does not show the model after the click' {
        Mock Wait-ForModelChoice { [pscustomobject]@{ Element = 'other'; Name = 'Model X' } }
        { Select-PreferredModel -Process $null -Target 'Model X' } | Should -Throw "*did not confirm the model 'Model X'*Auto, model picker*"
    }
    It 'does not touch the menu when the model is already selected' {
        Mock Find-ModelSwitcher { [pscustomobject]@{ Element = 'switcher'; Name = 'Model X, model picker' } }
        Select-PreferredModel -Process $null -Target 'model x' | Should -Be 'Model X, model picker'
        Should -Invoke Click-UiElement -Times 0
    }
}

Describe 'Prompt input without foreground' {
    BeforeAll {
        function New-FakeInput { $o = [pscustomobject]@{ Name = 'input' }; $o | Add-Member -MemberType ScriptMethod -Name SetFocus -Value {}; $o }
        $script:Prompt = "# Heading`n`n- first point`n1. Section one`nText with [SYN-1] and 3 numbers."
    }
    BeforeEach {
        Mock Find-CopilotInput { New-FakeInput }
        Mock Get-RenderWidget { [IntPtr]42 }
        Mock Send-TextMessages {}
        Mock Clear-CopilotInput {}
        Mock Send-KeyMessage {}
        Mock Find-SendButton { 'send-button' }
        $script:clicked = @()
        Mock Click-UiElement { $script:clicked += $Element }
    }
    It 'compares the typed text without white space and Markdown marks' {
        Get-ComparableText "a  b`n`nc" -Strict | Should -Be 'abc'
        Get-ComparableText "## Head`n- item`n2. Step" | Should -Be (Get-ComparableText "Head item Step")
        Get-ComparableText 'Case [SYN-1] costs 3.' | Should -Not -Be (Get-ComparableText 'Case [SYN-2] costs 3.')
    }
    It 'types the prompt and sends it with the send button when the input holds it' {
        Mock Get-InputValue { if ($script:typed) { "Heading`nfirst point`nSection one`nText with [SYN-1] and 3 numbers." } else { '' } }
        Mock Send-TextMessages { $script:typed = $true }
        $script:typed = $false
        $method = Send-CopilotPrompt -Process ([pscustomobject]@{ MainWindowHandle = [IntPtr]1 }) -Text $script:Prompt -TimeoutSeconds 5
        $method | Should -Match '^WindowMessages.*\+SendButton$'
        $script:clicked | Should -Be @('send-button')
    }
    It 'does not send when the input does not hold the whole prompt' {
        Mock Get-InputValue { 'Heading first point' }
        { Send-CopilotPrompt -Process ([pscustomobject]@{ MainWindowHandle = [IntPtr]1 }) -Text $script:Prompt -TimeoutSeconds 2 -StallSeconds 1 } |
            Should -Throw '*does not hold the prompt*Nothing was sent*'
        $script:clicked.Count | Should -Be 0
        Should -Invoke Send-KeyMessage -Times 0
    }
}

Describe 'Reading the answer from the UI Automation tree' {
    BeforeAll {
        function N([string]$Type, [string]$Name = '', [object[]]$Children = @(), [string]$Id = '') { [pscustomobject]@{ Type = $Type; Name = $Name; Id = $Id; Children = $Children } }
        # Shape of the Copilot answer as UI Automation shows it (30 October 2026).
        $script:Answer = N 'Group' 'Copilot said: ...' @(
            (N 'Text' 'Copilot said:' -Id 'copilot-message-_r_1_-title'),
            (N 'Group' '' @(
                    (N 'Group' '' @(
                            (N 'Text' 'Complaint'),
                            (N 'Text' 'Case [SYN-1] was received.'),
                            (N 'Button' 'Citation: shipment-log.csv'),
                            (N 'List' '' @((N 'ListItem' '' @((N 'Text' 'first point'))), (N 'ListItem' '' @((N 'Text' 'second point'))))),
                            (N 'List' '' @((N 'ListItem' 'Conclusion'))),
                            (N 'Text' 'Not justified.')
                        )),
                    (N 'Button' 'shipment-log'),
                    (N 'Text' 'From '),
                    (N 'Group' 'From A. Person' @((N 'Text' 'A. Person')))
                )),
            (N 'ToolBar' '' @((N 'Button' 'Copy response')))
        )
    }
    It 'returns headings and paragraphs as lines, without citations, file card, author line or toolbar' {
        $lines = @(Get-UiTextLines (Get-AnswerBodyNode $script:Answer))
        $lines -join '|' | Should -Be 'Complaint|Case [SYN-1] was received.|- first point|- second point|- Conclusion|Not justified.'
    }
    It 'recognises a text that is the prompt' {
        Test-PromptEcho "### Complaint record (from the complaint list)`n| a |" 'prompt' | Should -BeTrue
        Test-PromptEcho 'the prompt' 'the prompt' | Should -BeTrue
        Test-PromptEcho 'Complaint report text' 'the prompt' | Should -BeFalse
    }
}

Describe 'Round-trip script on Windows' -Tag 'Windows' -Skip:($env:OS -ne 'Windows_NT') {
    BeforeAll {
        $script:Script = Join-Path $script:Root 'src/Copilot/Invoke-CopilotRoundtrip.ps1'
        $script:Prompt = Join-Path $TestDrive 'prompt.txt'
        Set-Content -LiteralPath $script:Prompt -Value 'test prompt'
        function Invoke-Roundtrip([string[]]$Extra) {
            $status = Join-Path $TestDrive ("status-{0}.json" -f [guid]::NewGuid().ToString('N'))
            $output = Join-Path $TestDrive 'response.md'
            $null = & powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File $script:Script -PromptFile $script:Prompt `
                -OutputFile $output -StatusFile $status -StartDelaySeconds 0 @Extra
            [pscustomobject]@{ Exit = $LASTEXITCODE; Status = (Get-Content -Raw $status | ConvertFrom-Json); Output = $output }
        }
    }
    It 'aborts before touching any window when no model is configured' {
        $run = Invoke-Roundtrip @()
        $run.Exit | Should -Be 1
        $run.Status.error | Should -Match 'No model configured'
        $run.Output | Should -Not -Exist
    }
    It 'aborts when the Copilot window is not open' {
        $run = Invoke-Roundtrip @('-PreferredModel', 'Model X', '-WindowTitle', 'No such window 7f3a')
        $run.Exit | Should -Be 1
        $run.Status.error | Should -Match "No such window 7f3a|input desktop"
        $run.Output | Should -Not -Exist
    }
}
