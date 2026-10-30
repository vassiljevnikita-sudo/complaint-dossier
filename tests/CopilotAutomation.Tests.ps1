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
