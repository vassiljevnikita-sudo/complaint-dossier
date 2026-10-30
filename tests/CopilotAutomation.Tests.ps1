# Pester 5 tests for the Copilot UI Automation logic.
# The model selection runs against mocks, so these tests run on any OS.
# The tests tagged 'Windows' start the real round-trip script in Windows
# PowerShell 5.1 and check that it aborts cleanly when no Copilot window exists.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    . (Join-Path $script:Root 'src/Copilot/CopilotAutomation.ps1')
}

Describe 'UI name matching' {
    It 'ignores case, spaces and punctuation except dots' {
        Test-UiNameExact 'GPT-5 Think Deeper' 'gpt5 think deeper' | Should -BeTrue
        Test-UiNameExact 'Model 4.1' 'Model 41' | Should -BeFalse
        Normalize-UiName '' | Should -Be ''
    }
}

Describe 'Select-PreferredModel' {
    BeforeEach {
        Mock Activate-CopilotWindow {}
        Mock Start-Sleep {}
        Mock Write-Host {}
        Mock Move-MouseToUiElement {}
        Mock Click-UiElement {}
        Mock Find-ModelSwitcher { [pscustomobject]@{ Element = 'switcher'; Name = 'Auto' } }
    }

    It 'refuses to run without a configured model' {
        { Select-PreferredModel -Process $null -Target '' } | Should -Throw '*No model configured*'
        Should -Invoke Click-UiElement -Times 0
    }
    It 'aborts when the menu does not offer the model and clicks nothing else' {
        Mock Wait-ForNamedUiElement { $null }
        { Select-PreferredModel -Process $null -Target 'Model X' } | Should -Throw '*Configured model not found in the Copilot menu: Model X*'
        Should -Invoke Click-UiElement -Times 1 -Exactly -ParameterFilter { $Element -eq 'switcher' }
    }
    It 'aborts when the configured model group is missing' {
        Mock Wait-ForNamedUiElement { $null }
        { Select-PreferredModel -Process $null -Target 'Model X' -Group 'More models' } | Should -Throw '*model group not found*More models*'
    }
    It 'opens the group submenu and clicks the configured model' {
        $script:hovered = $false
        Mock Move-MouseToUiElement { $script:hovered = $true }
        Mock Wait-ForNamedUiElement {
            if ($Target -eq 'More models') { return [pscustomobject]@{ Element = 'group'; Name = 'More models' } }
            if ($Target -eq 'Model X' -and $script:hovered) { return [pscustomobject]@{ Element = 'model-x'; Name = 'Model X' } }
            $null
        }
        Select-PreferredModel -Process $null -Target 'Model X' -Group 'More models' | Should -Be 'Auto'
        Should -Invoke Click-UiElement -Times 1 -Exactly -ParameterFilter { $Element -eq 'model-x' }
    }
    It 'does not touch the menu when the model is already selected' {
        Mock Find-ModelSwitcher { [pscustomobject]@{ Element = 'switcher'; Name = 'Model X' } }
        Select-PreferredModel -Process $null -Target 'model x' | Should -Be 'Model X'
        Should -Invoke Click-UiElement -Times 0
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
