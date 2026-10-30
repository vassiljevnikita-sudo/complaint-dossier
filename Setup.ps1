<#
.SYNOPSIS
  First command: prepares Complaint Dossier on this computer and checks that
  your Excel list and your Word templates fit the configuration.

.DESCRIPTION
  1. Checks PowerShell, Windows and the Microsoft 365 Copilot app.
  2. Creates config.json from config.example.json if it does not exist yet and
     asks a few questions: your Excel list, the worksheet, which column holds
     what, your Word templates, the period and the Copilot model.
  3. Checks everything: the columns named in config.json exist in the list,
     every placeholder in the Word templates is known, every AI section has a
     place in the template. It prints what is missing and how to fix it.
  Run it again after every change. When it says "Ready", start Run.cmd.

.EXAMPLE
  .\Setup.cmd
.EXAMPLE
  .\Setup.cmd -Reconfigure
  Asks the questions again for an existing config.json.
.EXAMPLE
  ./Setup.ps1 -Demo -NonInteractive
  Checks the example configuration only (used in CI).
#>
[CmdletBinding()]
param(
    # Default: config.json, or config.example.json with -Demo.
    [string]$ConfigPath,
    # Check the example configuration for demo mode (no Copilot needed).
    [switch]$Demo,
    # Ask the questions again although config.json exists.
    [switch]$Reconfigure,
    # Never ask; only create and check.
    [switch]$NonInteractive,
    # Answers for the questions, in order (for tests). Empty string = Enter.
    [string[]]$Answers
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = $PSScriptRoot
Import-Module (Join-Path (Join-Path $root 'src') 'ComplaintDossier.psm1') -Force

$mode = 'Live'
if ($Demo) { $mode = 'Demo' }
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $root 'config.json'
    if ($Demo) { $ConfigPath = Join-Path $root 'config.example.json' }
}
$ConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)
$example = Join-Path $root 'config.example.json'

$script:AnswerQueue = $null
if ($null -ne $Answers) { $script:AnswerQueue = New-Object System.Collections.Generic.Queue[string]; foreach ($a in $Answers) { $script:AnswerQueue.Enqueue([string]$a) } }
$interactive = ($null -ne $script:AnswerQueue) -or (-not $NonInteractive -and [Environment]::UserInteractive -and -not $env:CI)

function Say([string]$Key, [object[]]$Arguments = @()) { Write-Host (Get-DossierText $Key -f $Arguments) }
function Ask([string]$Question, [string]$Default = '') {
    $prompt = $Question
    if ($Default) { $prompt += " [$Default]" }
    if ($null -ne $script:AnswerQueue) {
        $answer = ''
        if ($script:AnswerQueue.Count) { $answer = $script:AnswerQueue.Dequeue() }
        Write-Host "$prompt`: $answer"
    } else { $answer = Read-Host $prompt }
    $answer = ([string]$answer).Trim().Trim('"').Trim("'").Trim()
    if (-not $answer) { return $Default }
    $answer
}
function Show-Finding($f) {
    $mark = @{ ok = 'ok   '; info = 'note '; warn = 'WARN '; error = 'FIX  ' }[$f.Level]
    $color = @{ ok = 'Green'; info = 'Gray'; warn = 'Yellow'; error = 'Red' }[$f.Level]
    Write-Host "   $mark " -ForegroundColor $color -NoNewline
    Write-Host $f.Text
    if ($f.Fix) { Write-Host "         $($f.Fix)" -ForegroundColor DarkGray }
}
function ConvertTo-StoredPath([string]$Path) {
    # Paths inside the project are stored relative to it, others in full.
    $full = [IO.Path]::GetFullPath($Path)
    $rootFull = [IO.Path]::GetFullPath($root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if ($full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { return $full.Substring($rootFull.Length).Replace('\', '/') }
    $full
}
function Resolve-Answer([string]$Path) {
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    [IO.Path]::GetFullPath([IO.Path]::Combine($root, $Path))
}
function Set-Value($Object, [string]$Name, $Value) {
    if ($Object.PSObject.Properties[$Name]) { $Object.$Name = $Value } else { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
}

# Language: from the configuration if there is one, else from Windows.
Set-DossierLanguage (Get-DefaultLanguage)
if (Test-Path -LiteralPath $ConfigPath -PathType Leaf) {
    try { Set-DossierLanguage ([string](Read-JsonFile $ConfigPath).Language) } catch { }
}

Write-Host ''
Say 'SetupTitle' @((Get-DossierVersion))

# ---- 1. Computer
$script:Step = 0
function Step([string]$Key) { $script:Step++; Write-Host ''; Write-Host ("{0}. {1}" -f $script:Step, (Get-DossierText $Key)) }
Step 'SetupStepComputer'
$computer = New-Object System.Collections.Generic.List[object]
if ($env:OS -eq 'Windows_NT') {
    # Files from a downloaded zip are blocked by Windows; unblock this folder.
    try { Get-ChildItem -LiteralPath $root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue } catch { }
}
$computer.Add((New-Finding ok 'computer' (Get-DossierText 'PowerShellOk' -f $PSVersionTable.PSVersion.ToString(), $PSVersionTable.PSEdition) ''))
if ($env:OS -ne 'Windows_NT') {
    $level = 'error'
    if ($Demo) { $level = 'info' }
    $computer.Add((New-Finding $level 'computer' (Get-DossierText 'NotWindows') ''))
} elseif (-not $Demo) {
    $package = $null
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') { Import-Module Appx -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop }
        $package = @(Get-AppxPackage -Name 'Microsoft.MicrosoftOfficeHub' -ErrorAction Stop)[0]
    } catch { $package = $null }
    if ($package) { $computer.Add((New-Finding ok 'computer' (Get-DossierText 'CopilotAppOk' -f $package.Version) '')) }
    else { $computer.Add((New-Finding warn 'computer' (Get-DossierText 'CopilotAppMissing') (Get-DossierText 'FixCopilotApp'))) }
}
foreach ($f in $computer) { Show-Finding $f }

# ---- 2. Settings file
Step 'SetupStepConfig'
$created = $false
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    if ($Demo) { throw "Configuration file not found: $ConfigPath" }
    Copy-Item -LiteralPath $example -Destination $ConfigPath
    $created = $true
    Show-Finding (New-Finding ok 'config' (Get-DossierText 'ConfigCreated' -f (ConvertTo-StoredPath $ConfigPath)) '')
} else {
    Show-Finding (New-Finding ok 'config' (Get-DossierText 'ConfigFound' -f (ConvertTo-StoredPath $ConfigPath)) '')
}

if (($created -or $Reconfigure) -and $interactive -and -not $Demo) {
    $cfg = Read-JsonFile $ConfigPath
    Write-Host ''
    Say 'QIntro'
    $lang = Ask (Get-DossierText 'QLanguage') (Get-DossierLanguage)
    if (@('de', 'en') -contains $lang.ToLowerInvariant()) { Set-Value $cfg 'Language' $lang.ToLowerInvariant(); Set-DossierLanguage $lang }

    # Excel list
    $workbook = Resolve-Answer ([string]$cfg.Workbook.Path)
    for ($try = 0; $try -lt 3; $try++) {
        $answer = Ask (Get-DossierText 'QWorkbook') (ConvertTo-StoredPath $workbook)
        $candidate = Resolve-Answer $answer
        if ((Test-Path -LiteralPath $candidate -PathType Leaf) -and [IO.Path]::GetExtension($candidate).ToLowerInvariant() -in '.xlsx', '.xlsm') { $workbook = $candidate; break }
        Say 'QNotFound' @($candidate)
    }
    $cfg.Workbook.Path = ConvertTo-StoredPath $workbook
    if (Test-Path -LiteralPath $workbook -PathType Leaf) {
        $sheets = @(Get-XlsxSheetNames $workbook)
        $sheet = [string]$cfg.Workbook.Sheet
        if ($sheets -notcontains $sheet) { $sheet = $sheets[0] }
        if ($sheets.Count -gt 1) {
            for ($i = 0; $i -lt $sheets.Count; $i++) { Write-Host ("     {0}  {1}" -f ($i + 1), $sheets[$i]) }
            $answer = Ask (Get-DossierText 'QSheet') ([string]($sheets.IndexOf($sheet) + 1))
            $index = 0
            if ([int]::TryParse($answer, [ref]$index) -and $index -ge 1 -and $index -le $sheets.Count) { $sheet = $sheets[$index - 1] }
            elseif ($sheets -contains $answer) { $sheet = $answer }
        }
        $cfg.Workbook.Sheet = $sheet
        $rows = @(Read-XlsxRows -Path $workbook -Sheet $sheet)
        $guess = Find-XlsxHeaderRow $rows
        if ($rows.Count) {
            $header = Get-XlsxHeader $rows $guess
            Say 'QHeaderShown' @($guess, ((@($header.Keys)) -join ' | '))
            $answer = Ask (Get-DossierText 'QHeaderRow') ([string]$guess)
            $row = 0
            if ([int]::TryParse($answer, [ref]$row) -and $row -ge 1) { $guess = $row }
        }
        Set-Value $cfg.Workbook 'HeaderRow' $guess
        $header = Get-XlsxHeader $rows $guess
        $names = @($header.Keys)
        $listed = $false
        $pick = {
            param([string]$Label, [string]$Current, [bool]$Required)
            if ($Current -and $names -contains $Current) { return $Current }
            if (-not $listed) {
                Say 'QColumnsIntro'
                for ($i = 0; $i -lt $names.Count; $i++) { Write-Host ("     {0,2}  {1}" -f ($i + 1), $names[$i]) }
                Set-Variable -Name listed -Value $true -Scope 1
            }
            for ($try = 0; $try -lt 3; $try++) {
                $q = Get-DossierText 'QColumn' -f $Label
                if (-not $Required) { $q += ' ' + (Get-DossierText 'QColumnOptional') }
                $answer = Ask $q ''
                $index = 0
                if ([int]::TryParse($answer, [ref]$index) -and $index -ge 1 -and $index -le $names.Count) { return $names[$index - 1] }
                if ($names -contains $answer) { return $answer }
                if (-not $answer -and -not $Required) { return '' }
            }
            $Current
        }
        foreach ($field in @('CaseId', 'ReceivedDate', 'Category', 'Description', 'Product', 'ProductGroup', 'Batch', 'Severity', 'Status', 'SampleReceived', 'ReportType')) {
            $current = [string](Get-ConfigValue $cfg.Workbook.Columns $field)
            $value = & $pick (Get-DossierText ('Role' + $field)) $current ($field -in 'CaseId', 'ReceivedDate')
            Set-Value $cfg.Workbook.Columns $field $value
        }
        foreach ($p in @($cfg.Template.Fields.PSObject.Properties)) {
            $value = & $pick ("{{$($p.Name)}}") ([string]$p.Value) $false
            $cfg.Template.Fields.($p.Name) = $value
        }
    }

    # Word templates
    foreach ($type in 'Short', 'Long') {
        $current = Resolve-Answer ([string]$cfg.Reports.$type.TemplatePath)
        for ($try = 0; $try -lt 3; $try++) {
            $answer = Ask (Get-DossierText ('QTemplate' + $type)) (ConvertTo-StoredPath $current)
            $candidate = Resolve-Answer $answer
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $current = $candidate; break }
            Say 'QNotFound' @($candidate)
        }
        $cfg.Reports.$type.TemplatePath = ConvertTo-StoredPath $current
    }

    # Period, case folders, model
    foreach ($key in 'From', 'To') {
        $answer = Ask (Get-DossierText ('QPeriod' + $key)) ([string]$cfg.Selection.$key)
        if ($answer -eq '-') { $answer = '' }
        $cfg.Selection.$key = $answer
    }
    $answer = Ask (Get-DossierText 'QCaseFolders') ([string]$cfg.CaseFolders.Path)
    if ($answer -eq '-') { $cfg.CaseFolders.Path = '' } else { $cfg.CaseFolders.Path = ConvertTo-StoredPath (Resolve-Answer $answer) }
    $cfg.Copilot.PreferredModel = Ask (Get-DossierText 'QModel') ([string]$cfg.Copilot.PreferredModel)

    Write-JsonFile $ConfigPath $cfg
    Say 'ConfigSaved' @((ConvertTo-StoredPath $ConfigPath))
}

# ---- 3. Checks
$config = Read-DossierConfig -Path $ConfigPath -Root $root -Mode $mode
$findings = @($computer.ToArray()) + @(Test-DossierSetup -Config $config)
if ($mode -eq 'Live' -and $env:OS -eq 'Windows_NT') {
    if (Test-CopilotWindowOpen ([string]$config.Settings.Copilot.WindowTitle)) { $findings += New-Finding ok 'copilot' (Get-DossierText 'CopilotOpen') '' }
    else { $findings += New-Finding info 'copilot' (Get-DossierText 'CopilotClosed') '' }
}
foreach ($section in 'config', 'excel', 'templates', 'copilot', 'output') {
    $items = @($findings | Where-Object { $_.Section -eq $section })
    if (-not $items.Count) { continue }
    # Findings about other settings belong to step 2, the settings file.
    if ($section -ne 'config') { Step ('SetupStep_' + $section) }
    foreach ($f in $items) { Show-Finding $f }
}

$errors = @($findings | Where-Object { $_.Level -eq 'error' })
$warnings = @($findings | Where-Object { $_.Level -eq 'warn' })
Write-Host ''
if ($errors.Count) {
    Write-Host (Get-DossierText 'SetupNotReady' -f $errors.Count) -ForegroundColor Red
    Say 'SetupAgain'
    exit 1
}
if ($Demo) { Write-Host (Get-DossierText 'SetupReadyDemo' -f $warnings.Count) -ForegroundColor Green }
else { Write-Host (Get-DossierText 'SetupReady' -f $warnings.Count) -ForegroundColor Green }
exit 0
