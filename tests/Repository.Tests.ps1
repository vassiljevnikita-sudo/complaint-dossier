# Pester 5 checks on the repository content itself.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:Files = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File | Where-Object {
            $relative = $_.FullName.Substring($script:Root.Length + 1).Replace('\', '/')
            $relative -notmatch '^(\.git|out)/' -and $relative -ne 'config.json'
        })
    function Get-Relative($File) { $File.FullName.Substring($script:Root.Length + 1).Replace('\', '/') }
}

Describe 'Repository' {
    It 'contains no machine-specific absolute paths' {
        # Built from parts so this file does not match itself.
        $pattern = '(?i)\b[A-Z]:\\(' + 'Users|Daten|Data)\\|/(' + 'home|mnt|Users)/[a-z]'
        $hits = foreach ($file in $script:Files | Where-Object { $_.Extension -in '.ps1', '.psm1', '.json', '.md', '.txt', '.yml' }) {
            if ((Get-Content -Raw -LiteralPath $file.FullName) -match $pattern) { Get-Relative $file }
        }
        $hits | Should -BeNullOrEmpty
    }
    It 'keeps every script ASCII-only, because Windows PowerShell 5.1 reads BOM-less files as ANSI' {
        $hits = foreach ($file in $script:Files | Where-Object { $_.Extension -in '.ps1', '.psm1' }) {
            if ([IO.File]::ReadAllBytes($file.FullName) | Where-Object { $_ -gt 127 } | Select-Object -First 1) { Get-Relative $file }
        }
        $hits | Should -BeNullOrEmpty
    }
    It 'marks every sample file as synthetic in its first line' {
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:Root 'samples') -Recurse -File -Include *.txt, *.md) {
            (Get-Content -LiteralPath $file.FullName -TotalCount 1) | Should -Match 'SYNTHETIC' -Because (Get-Relative $file)
        }
    }
    It 'ships an example configuration without a preselected model' {
        $config = Get-Content -Raw (Join-Path $script:Root 'config.example.json') | ConvertFrom-Json
        $config.Copilot.PreferredModel | Should -Be ''
    }
    It 'keeps local configuration and run output out of git' {
        $ignore = Get-Content (Join-Path $script:Root '.gitignore')
        $ignore | Should -Contain '/config.json'
        $ignore | Should -Contain '/out/'
    }
}
