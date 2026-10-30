# Complaint Dossier core functions.
# Generates one draft complaint report per complaint: complaint list (Excel) ->
# selection by period and status -> per-complaint prompt -> Microsoft 365 Copilot
# (or a prewritten demo answer) -> the user's own Word template, filled ->
# review list and audit log. The review happens afterwards, in Word.
# Runs on PowerShell 7 (Windows, Linux, macOS) and on Windows PowerShell 5.1.
# Keep this file ASCII-only: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.

Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:ToolVersion = '0.4.0'
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:FixedZipTime = [DateTimeOffset]::new(2026, 1, 1, 0, 0, 0, [TimeSpan]::Zero)

# Headings of the data part of every prompt. A response that contains one of
# them is a copy of the prompt, not an answer.
$script:PromptHeadings = @(
    '### Complaint record (from the complaint list)',
    '### Similar complaints (calculated by the tool)',
    '### Case documents',
    '### Attached files'
)

# Fields of the complaint list the tool itself works with. Workbook.Columns maps
# them to the column names of the user's list. CaseId and ReceivedDate are required.
$script:CaseFields = @('CaseId', 'ReceivedDate', 'Category', 'Description', 'Product', 'ProductGroup', 'Batch', 'Severity', 'Status', 'SampleReceived', 'ReportType')
$script:RequiredCaseFields = @('CaseId', 'ReceivedDate')
$script:ReportTypes = @('Short', 'Long')

# Placeholders the tool fills itself. Template.Fields and the AI sections of
# Reports.<type>.Sections must not use these names.
$script:ToolPlaceholders = [ordered]@{
    DraftNotice       = 'inline'
    ReportTitle       = 'inline'
    ReportType        = 'inline'
    GeneratedAt       = 'inline'
    Model             = 'inline'
    RunId             = 'inline'
    FrequencyLevel    = 'inline'
    CheckSummary      = 'block'
    SimilarComplaints = 'block'
    CaseDocuments     = 'block'
    ReportText        = 'block'
}

. (Join-Path $PSScriptRoot 'WordTemplate.ps1')

function Get-DossierVersion { $script:ToolVersion }

# ---------------------------------------------------------------- messages (English / German)

$script:Messages = $null
$script:Language = 'en'

function Get-DefaultLanguage {
    try { if ((Get-UICulture).TwoLetterISOLanguageName -eq 'de') { return 'de' } } catch { }
    'en'
}

function Set-DossierLanguage {
    <# Language of console messages, check names in the report and the review list: 'en' or 'de'. #>
    param([string]$Language)
    if ($Language -and @('en', 'de') -contains $Language.ToLowerInvariant()) { $script:Language = $Language.ToLowerInvariant() }
}

function Get-DossierLanguage { $script:Language }

function Get-DossierText {
    <# Message text in the current language, with {0} {1} ... filled from -f. #>
    param([Parameter(Mandatory)][string]$Key, [object[]]$f = @())
    if ($null -eq $script:Messages) {
        $script:Messages = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'messages.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    }
    $table = $script:Messages.($script:Language)
    $text = $null
    if ($table.PSObject.Properties[$Key]) { $text = [string]$table.$Key }
    elseif ($script:Messages.en.PSObject.Properties[$Key]) { $text = [string]$script:Messages.en.$Key }
    else { throw "Message key missing: $Key" }
    if (@($f).Count) { $text = [string]::Format([Globalization.CultureInfo]::InvariantCulture, $text, [object[]]$f) }
    # [[Name]] in a message stands for the placeholder {{Name}}, so braces need no escaping.
    $text.Replace('[[', '{{').Replace(']]', '}}')
}
Set-Alias -Name T -Value Get-DossierText -Scope Script

# ---------------------------------------------------------------- files, hashes

function Write-Utf8File {
    param([Parameter(Mandatory)][string]$Path, [AllowEmptyString()][string]$Text)
    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    [IO.File]::WriteAllText($Path, $Text, $script:Utf8NoBom)
}

function Read-Utf8File {
    param([Parameter(Mandatory)][string]$Path)
    [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
}

function Get-Sha256 {
    param([Parameter(Mandatory)][string]$Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-TextSha256 {
    param([AllowEmptyString()][string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash($script:Utf8NoBom.GetBytes($Text))
        ([BitConverter]::ToString($bytes)).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
}

function Get-UtcStamp {
    # One timestamp format everywhere, so audit.json survives a read/write cycle unchanged.
    [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function Restore-JsonDate {
    # PowerShell 7.0-7.4 turn ISO date strings into [DateTime] when reading JSON.
    # Turn them back into the original text form.
    param($Value)
    if ($Value -is [DateTime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($Value -is [Management.Automation.PSCustomObject]) {
        foreach ($p in @($Value.PSObject.Properties)) { $p.Value = Restore-JsonDate $p.Value }
        return $Value
    }
    if ($Value -is [Array]) { return , @($Value | ForEach-Object { Restore-JsonDate $_ }) }
    $Value
}

function Read-JsonFile {
    <# Reads JSON and keeps date strings as strings on every PowerShell version. #>
    param([Parameter(Mandatory)][string]$Path)
    $text = Read-Utf8File $Path
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { return $text | ConvertFrom-Json -DateKind String }
    $object = $text | ConvertFrom-Json
    if ($PSVersionTable.PSVersion.Major -ge 6) { $object = Restore-JsonDate $object }
    $object
}

function Resolve-DossierPath {
    # Relative paths in the config resolve against the project root.
    param([Parameter(Mandatory)][string]$Root, [AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    [IO.Path]::GetFullPath([IO.Path]::Combine($Root, $Path))
}

function Get-RelativeDisplayPath {
    # Paths inside the project are logged relative to it, so audit logs
    # do not carry user or machine specific folders.
    param([string]$Root, [string]$Path, [switch]$FullPathOutside)
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($rootFull.Length).Replace('\', '/')
    }
    if ($FullPathOutside) { return $full }
    [IO.Path]::GetFileName($full)
}


function ConvertTo-DossierJson {
    <#
      JSON with two-space indent, LF line ends and non-ASCII characters kept as
      they are. The same object gives the same bytes on Windows PowerShell 5.1
      and PowerShell 7, so hashes of JSON files do not depend on the version.
    #>
    param($Value, [int]$Indent = 0)
    $pad = '  ' * $Indent
    $inner = '  ' * ($Indent + 1)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte]) { return ([string]$Value) }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) { return $Value.ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [DateTime]) { $Value = $Value.ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($Value -is [string] -or $Value -is [char]) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('"')
        foreach ($ch in ([string]$Value).ToCharArray()) {
            switch ($ch) {
                '"' { [void]$sb.Append('\"') }
                '\' { [void]$sb.Append('\\') }
                "`n" { [void]$sb.Append('\n') }
                "`r" { [void]$sb.Append('\r') }
                "`t" { [void]$sb.Append('\t') }
                default {
                    if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } else { [void]$sb.Append($ch) }
                }
            }
        }
        [void]$sb.Append('"')
        return $sb.ToString()
    }
    if ($Value -is [Collections.IDictionary]) {
        $keys = @($Value.Keys)
        if (-not $keys.Count) { return '{}' }
        $items = foreach ($k in $keys) { $inner + (ConvertTo-DossierJson ([string]$k)) + ': ' + (ConvertTo-DossierJson $Value[$k] ($Indent + 1)) }
        return "{`n" + ($items -join ",`n") + "`n$pad}"
    }
    if ($Value -is [Management.Automation.PSCustomObject]) {
        $props = @($Value.PSObject.Properties)
        if (-not $props.Count) { return '{}' }
        $items = foreach ($p in $props) { $inner + (ConvertTo-DossierJson $p.Name) + ': ' + (ConvertTo-DossierJson $p.Value ($Indent + 1)) }
        return "{`n" + ($items -join ",`n") + "`n$pad}"
    }
    if ($Value -is [Collections.IEnumerable]) {
        $list = @($Value)
        if (-not $list.Count) { return '[]' }
        $items = foreach ($v in $list) { $inner + (ConvertTo-DossierJson $v ($Indent + 1)) }
        return "[`n" + ($items -join ",`n") + "`n$pad]"
    }
    ConvertTo-DossierJson ([string]$Value) $Indent
}

function Write-JsonFile {
    param([Parameter(Mandatory)][string]$Path, $Value)
    Write-Utf8File $Path ((ConvertTo-DossierJson $Value) + "`n")
}

function Get-SafeFileName {
    <# A complaint number as a file or folder name: characters Windows does not allow become '_'. #>
    param([Parameter(Mandatory)][string]$Name)
    $safe = [regex]::Replace($Name.Trim(), '[\\/:*?"<>|\x00-\x1f]', '_').TrimEnd('.', ' ')
    if (-not $safe) { $safe = '_' }
    $safe
}

# ---------------------------------------------------------------- configuration

function Get-ConfigValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } else { return $null } }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    $property.Value
}

function Get-ConfigString {
    param($Object, [string]$Name)
    $value = Get-ConfigValue $Object $Name
    if ($null -eq $value) { return '' }
    ([string]$value).Trim()
}

function Get-ConfigList {
    param($Object, [string]$Name)
    @(@(Get-ConfigValue $Object $Name) | Where-Object { $null -ne $_ -and [string]$_ -ne '' })
}

function Read-DossierConfig {
    <#
      Reads the JSON configuration and resolves every path in it against the
      project folder. Stops only when the file is missing or not valid JSON;
      Test-DossierConfig and Test-DossierSetup check the content.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root,
        [ValidateSet('Demo', 'Live')][string]$Mode = 'Demo',
        # Replaces Output.Path, for example for CI runs.
        [string]$OutputPath = ''
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw (T 'ConfigMissing' -f $Path) }
    try { $config = Read-JsonFile $Path }
    catch { throw (T 'ConfigInvalid' -f $Path, $_.Exception.Message) }
    Set-DossierLanguage (Get-ConfigString $config 'Language')

    $templates = [ordered]@{}
    $prompts = [ordered]@{}
    foreach ($type in $script:ReportTypes) {
        $settings = Get-ConfigValue (Get-ConfigValue $config 'Reports') $type
        $templates[$type] = Resolve-DossierPath $Root (Get-ConfigString $settings 'TemplatePath')
        $prompts[$type] = Resolve-DossierPath $Root (Get-ConfigString $settings 'PromptPath')
    }
    $out = Get-ConfigString (Get-ConfigValue $config 'Output') 'Path'
    if (-not $out) { $out = 'out' }
    if ($OutputPath) { $output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath) }
    else { $output = Resolve-DossierPath $Root $out }
    $paths = [ordered]@{
        Config        = [IO.Path]::GetFullPath($Path)
        Workbook      = Resolve-DossierPath $Root (Get-ConfigString (Get-ConfigValue $config 'Workbook') 'Path')
        CaseFolders   = Resolve-DossierPath $Root (Get-ConfigString (Get-ConfigValue $config 'CaseFolders') 'Path')
        Templates     = [pscustomobject]$templates
        Prompts       = [pscustomobject]$prompts
        Output        = $output
        Reports       = Join-Path $output 'reports'
        Audit         = Join-Path $output 'audit'
        Runs          = Join-Path $output 'runs'
        ReviewList    = Join-Path $output 'review-list.xlsx'
        DemoResponses = Resolve-DossierPath $Root (Get-ConfigString (Get-ConfigValue $config 'Demo') 'ResponsePath')
        Upload        = Resolve-DossierPath $Root (Get-ConfigString (Get-ConfigValue (Get-ConfigValue $config 'Copilot') 'Upload') 'Path')
    }
    [pscustomobject]@{ Settings = $config; Paths = [pscustomobject]$paths; Root = $Root; Mode = $Mode }
}

function New-Finding {
    param([ValidateSet('ok', 'info', 'warn', 'error')][string]$Level, [string]$Section, [string]$Text, [string]$Fix = '')
    [pscustomobject]@{ Level = $Level; Section = $Section; Text = $Text; Fix = $Fix }
}

function Test-DateText {
    param([string]$Text)
    $parsed = [DateTime]::MinValue
    [DateTime]::TryParseExact($Text, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$parsed)
}

function Get-SectionSettings {
    <# Reports.<type>.Sections as objects with Heading and Placeholder. #>
    param($Settings)
    foreach ($s in @(Get-ConfigValue $Settings 'Sections')) {
        if ($null -eq $s) { continue }
        if ($s -is [string]) { [pscustomobject]@{ Heading = $s; Placeholder = '' } }
        else { [pscustomobject]@{ Heading = (Get-ConfigString $s 'Heading'); Placeholder = (Get-ConfigString $s 'Placeholder') } }
    }
}

function Test-DossierConfig {
    <# Checks the configuration values and the files they point to. Returns findings; does not read Excel or Word. #>
    param([Parameter(Mandatory)]$Config)
    $c = $Config.Settings
    $p = $Config.Paths
    $section = 'config'
    $wb = Get-ConfigValue $c 'Workbook'
    if (-not (Get-ConfigString $wb 'Path')) { New-Finding error 'excel' (T 'ValueMissing' -f 'Workbook.Path') (T 'FixWorkbookPath') }
    elseif (-not (Test-Path -LiteralPath $p.Workbook -PathType Leaf)) { New-Finding error 'excel' (T 'FileNotFound' -f $p.Workbook) (T 'FixWorkbookPath') }
    elseif ([IO.Path]::GetExtension($p.Workbook).ToLowerInvariant() -notin '.xlsx', '.xlsm') { New-Finding error 'excel' (T 'WorkbookNotXlsx' -f ([IO.Path]::GetFileName($p.Workbook))) (T 'FixSaveAsXlsx') }
    if (-not (Get-ConfigString $wb 'Sheet')) { New-Finding error 'excel' (T 'ValueMissing' -f 'Workbook.Sheet') '' }
    foreach ($field in $script:RequiredCaseFields) {
        if (-not (Get-ConfigString (Get-ConfigValue $wb 'Columns') $field)) { New-Finding error 'excel' (T 'ValueMissing' -f "Workbook.Columns.$field") (T ('Role' + $field)) }
    }

    $sel = Get-ConfigValue $c 'Selection'
    foreach ($key in 'From', 'To') {
        $value = Get-ConfigString $sel $key
        if ($value -and -not (Test-DateText $value)) { New-Finding error $section (T 'BadDateSetting' -f "Selection.$key", $value) '' }
    }
    $from = Get-ConfigString $sel 'From'
    $to = Get-ConfigString $sel 'To'
    if ($from -and $to -and (Test-DateText $from) -and (Test-DateText $to) -and $from -gt $to) { New-Finding error $section (T 'PeriodReversed' -f $from, $to) '' }

    if ($p.CaseFolders -and -not (Test-Path -LiteralPath $p.CaseFolders -PathType Container)) {
        New-Finding error $section (T 'FolderNotFound' -f 'CaseFolders.Path', $p.CaseFolders) (T 'FixCaseFolders')
    }

    $reports = Get-ConfigValue $c 'Reports'
    $default = Get-ConfigString $reports 'DefaultType'
    if ($default -and $script:ReportTypes -notcontains $default) { New-Finding error $section (T 'BadDefaultType' -f $default) '' }
    $fieldNames = @(@(Get-ConfigValue (Get-ConfigValue $c 'Template') 'Fields') | ForEach-Object { if ($_) { $_.PSObject.Properties.Name } })
    foreach ($name in $fieldNames) {
        if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { New-Finding error 'templates' (T 'BadPlaceholderName' -f $name, 'Template.Fields') (T 'FixPlaceholderName') }
        if ($script:ToolPlaceholders.Contains($name)) { New-Finding error 'templates' (T 'PlaceholderReserved' -f $name, 'Template.Fields') '' }
    }
    $markers = @()
    foreach ($type in $script:ReportTypes) {
        $s = Get-ConfigValue $reports $type
        if ($null -eq $s) { New-Finding error $section (T 'ValueMissing' -f "Reports.$type") ''; continue }
        foreach ($key in 'Title', 'Marker') {
            if (-not (Get-ConfigString $s $key)) { New-Finding error $section (T 'ValueMissing' -f "Reports.$type.$key") '' }
        }
        $markers += (Get-ConfigString $s 'Marker').ToLowerInvariant()
        $template = $p.Templates.$type
        if (-not (Get-ConfigString $s 'TemplatePath')) { New-Finding error 'templates' (T 'ValueMissing' -f "Reports.$type.TemplatePath") (T 'FixTemplatePath') }
        elseif (-not (Test-Path -LiteralPath $template -PathType Leaf)) { New-Finding error 'templates' (T 'FileNotFound' -f $template) (T 'FixTemplatePath') }
        elseif ([IO.Path]::GetExtension($template).ToLowerInvariant() -notin '.docx', '.dotx') { New-Finding error 'templates' (T 'TemplateNotDocx' -f ([IO.Path]::GetFileName($template))) (T 'FixSaveAsDocx') }
        if (-not (Get-ConfigString $s 'PromptPath')) { New-Finding error $section (T 'ValueMissing' -f "Reports.$type.PromptPath") '' }
        elseif (-not (Test-Path -LiteralPath $p.Prompts.$type -PathType Leaf)) { New-Finding error $section (T 'FileNotFound' -f $p.Prompts.$type) '' }
        $sections = @(Get-SectionSettings $s)
        if (-not $sections.Count) { New-Finding error $section (T 'ValueMissing' -f "Reports.$type.Sections") '' }
        $seen = @{}
        foreach ($sec in $sections) {
            if (-not $sec.Heading) { New-Finding error $section (T 'ValueMissing' -f "Reports.$type.Sections[].Heading") ''; continue }
            if (-not $sec.Placeholder) { continue }
            if ($sec.Placeholder -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { New-Finding error 'templates' (T 'BadPlaceholderName' -f $sec.Placeholder, "Reports.$type.Sections") (T 'FixPlaceholderName') }
            if ($script:ToolPlaceholders.Contains($sec.Placeholder)) { New-Finding error 'templates' (T 'PlaceholderReserved' -f $sec.Placeholder, "Reports.$type.Sections") '' }
            if ($fieldNames -contains $sec.Placeholder) { New-Finding error 'templates' (T 'PlaceholderTwice' -f $sec.Placeholder) '' }
            if ($seen.ContainsKey($sec.Placeholder)) { New-Finding error 'templates' (T 'PlaceholderTwice' -f $sec.Placeholder) '' }
            $seen[$sec.Placeholder] = $true
        }
    }
    if ($markers.Count -eq 2 -and $markers[0] -and $markers[0] -eq $markers[1]) { New-Finding error $section (T 'MarkersEqual') '' }

    if ($Config.Mode -eq 'Demo' -and -not (Test-Path -LiteralPath $p.DemoResponses -PathType Container)) {
        New-Finding error $section (T 'FolderNotFound' -f 'Demo.ResponsePath', $p.DemoResponses) ''
    }
    if ($Config.Mode -eq 'Live') {
        $cp = Get-ConfigValue $c 'Copilot'
        if (-not (Get-ConfigString $cp 'WindowTitle')) { New-Finding error 'copilot' (T 'ValueMissing' -f 'Copilot.WindowTitle') '' }
        if (-not (Get-ConfigString $cp 'PreferredModel')) { New-Finding error 'copilot' (T 'ModelMissing') (T 'FixModel') }
        else { New-Finding ok 'copilot' (T 'ModelSet' -f (Get-ConfigString $cp 'PreferredModel')) '' }
    }
}

function Import-DossierConfig {
    <# Reads the configuration and stops on the first configuration error. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root,
        [ValidateSet('Demo', 'Live')][string]$Mode = 'Demo',
        [string]$OutputPath = ''
    )
    $config = Read-DossierConfig -Path $Path -Root $Root -Mode $Mode -OutputPath $OutputPath
    $errors = @(Test-DossierConfig $config | Where-Object { $_.Level -eq 'error' })
    if ($errors.Count) {
        $text = $errors[0].Text
        if ($errors[0].Fix) { $text += ' ' + $errors[0].Fix }
        throw $text
    }
    $config
}

# ---------------------------------------------------------------- xlsx (OpenXML, no Excel needed)

function Get-ZipEntryText {
    param([IO.Compression.ZipArchive]$Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { return $null }
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function ConvertFrom-CellReference {
    # "C12" -> zero-based column index 2
    param([string]$Reference)
    $letters = ($Reference -replace '[0-9]', '').ToUpperInvariant()
    $index = 0
    foreach ($ch in $letters.ToCharArray()) { $index = $index * 26 + ([int]$ch - 64) }
    $index - 1
}

function ConvertTo-CellReference {
    # zero-based column index 2, row 12 -> "C12"
    param([int]$Column, [int]$Row)
    $name = ''
    $n = $Column + 1
    while ($n -gt 0) { $m = ($n - 1) % 26; $name = [string][char](65 + $m) + $name; $n = [int][math]::Floor(($n - $m) / 26) }
    "$name$Row"
}

function Test-DateNumberFormat {
    <# True when an Excel number format shows a date (d or y outside quotes and brackets). #>
    param([int]$Id, [string]$Code)
    if (@(14, 15, 16, 17, 22) + (27..36) + (50..58) -contains $Id) { return $true }
    if (-not $Code) { return $false }
    $plain = [regex]::Replace($Code, '"[^"]*"|\[[^\]]*\]|\\.', '')
    $plain -match '[dDyY]'
}

function Get-XlsxSheetNames {
    param([Parameter(Mandatory)][string]$Path)
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        $zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
        $workbook = [xml](Get-ZipEntryText $zip 'xl/workbook.xml')
        @($workbook.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') })
    } finally { $stream.Dispose() }
}

function Read-XlsxRows {
    <#
      Reads the cells of one worksheet: one entry per row with Number and Values
      (zero-based column index -> value). Pure OpenXML: works without Excel.
      Numbers are [double], cells with a date format [DateTime], text [string].
    #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Sheet)
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        $zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
        $workbook = [xml](Get-ZipEntryText $zip 'xl/workbook.xml')
        $sheetNode = @($workbook.SelectNodes("//*[local-name()='sheet']") | Where-Object { $_.GetAttribute('name') -eq $Sheet })
        if ($sheetNode.Count -ne 1) {
            $names = @($workbook.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') }) -join ', '
            throw (T 'SheetNotFound' -f $Sheet, ([IO.Path]::GetFileName($Path)), $names)
        }
        $relId = $sheetNode[0].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $rels = [xml](Get-ZipEntryText $zip 'xl/_rels/workbook.xml.rels')
        $target = @($rels.SelectNodes("//*[local-name()='Relationship']") | Where-Object { $_.GetAttribute('Id') -eq $relId })[0].GetAttribute('Target')
        if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }

        $shared = @()
        $sharedXml = Get-ZipEntryText $zip 'xl/sharedStrings.xml'
        if ($sharedXml) {
            $shared = @(([xml]$sharedXml).SelectNodes("//*[local-name()='si']") | ForEach-Object {
                    (@($_.SelectNodes(".//*[local-name()='t'][not(ancestor::*[local-name()='rPh'])]") | ForEach-Object { $_.InnerText }) -join '')
                })
        }
        # Cell styles that show a date: the number is an Excel date serial.
        $dateStyles = @{}
        $stylesXml = Get-ZipEntryText $zip 'xl/styles.xml'
        if ($stylesXml) {
            $styles = [xml]$stylesXml
            $codes = @{}
            foreach ($f in @($styles.SelectNodes("//*[local-name()='numFmts']/*[local-name()='numFmt']"))) { $codes[[int]$f.GetAttribute('numFmtId')] = $f.GetAttribute('formatCode') }
            $i = 0
            foreach ($xf in @($styles.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']"))) {
                $id = 0
                if ($xf.GetAttribute('numFmtId')) { $id = [int]$xf.GetAttribute('numFmtId') }
                if (Test-DateNumberFormat $id ([string]$codes[$id])) { $dateStyles[$i] = $true }
                $i++
            }
        }

        $sheetXml = [xml](Get-ZipEntryText $zip $target)
        foreach ($row in $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $values = @{}
            foreach ($cell in $row.SelectNodes("*[local-name()='c']")) {
                $col = ConvertFrom-CellReference $cell.GetAttribute('r')
                $type = $cell.GetAttribute('t')
                $v = $cell.SelectSingleNode("*[local-name()='v']")
                $value = $null
                switch ($type) {
                    's' { if ($v) { $value = [string]$shared[[int]$v.InnerText] } }
                    'inlineStr' { $value = (@($cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join '') }
                    'str' { if ($v) { $value = [string]$v.InnerText } }
                    'b' { if ($v) { $value = ($v.InnerText -eq '1') } }
                    'e' { $value = $null }
                    default {
                        if ($v -and $v.InnerText -ne '') {
                            $value = [double]::Parse($v.InnerText, [Globalization.CultureInfo]::InvariantCulture)
                            $style = $cell.GetAttribute('s')
                            if ($style -and $dateStyles.ContainsKey([int]$style) -and $value -gt 0 -and $value -lt 2958466) { $value = [DateTime]::FromOADate($value) }
                        }
                    }
                }
                $values[$col] = $value
            }
            [pscustomobject]@{ Number = [int]$row.GetAttribute('r'); Values = $values }
        }
    } finally { $stream.Dispose() }
}

function Get-XlsxHeader {
    <# Column names in the header row: zero-based column index -> name. #>
    param([Parameter(Mandatory)][object[]]$Rows, [int]$HeaderRow = 1)
    $header = @($Rows | Where-Object { $_.Number -eq $HeaderRow })
    $columns = [ordered]@{}
    if ($header.Count -ne 1) { return $columns }
    foreach ($key in @($header[0].Values.Keys | Sort-Object)) {
        $name = ([string]$header[0].Values[$key]).Trim()
        if ($name -and -not $columns.Contains($name)) { $columns[$name] = $key }
    }
    $columns
}

function Find-XlsxHeaderRow {
    <# Guesses the header row: the first of the top 20 rows with at least 3 text cells and no numbers or dates. #>
    param([Parameter(Mandatory)][object[]]$Rows)
    foreach ($row in @($Rows | Where-Object { $_.Number -le 20 } | Sort-Object Number)) {
        $values = @($row.Values.Values | Where-Object { $null -ne $_ -and [string]$_ -ne '' })
        $text = @($values | Where-Object { $_ -is [string] })
        if ($text.Count -ge 3 -and $text.Count -eq $values.Count) { return $row.Number }
    }
    1
}

function Read-XlsxSheet {
    <# Reads one worksheet into objects keyed by the header row. Empty rows are left out. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Sheet,
        [int]$HeaderRow = 1
    )
    $rows = @(Read-XlsxRows -Path $Path -Sheet $Sheet)
    $columns = Get-XlsxHeader $rows $HeaderRow
    if (-not $columns.Count) { throw (T 'HeaderRowEmpty' -f $HeaderRow, $Sheet) }
    foreach ($row in @($rows | Where-Object { $_.Number -gt $HeaderRow } | Sort-Object Number)) {
        $item = [ordered]@{ _Row = $row.Number }
        $empty = $true
        foreach ($name in $columns.Keys) {
            $value = $row.Values[$columns[$name]]
            if ($null -ne $value -and [string]$value -ne '') { $empty = $false }
            $item[$name] = $value
        }
        if (-not $empty) { [pscustomobject]$item }
    }
}

function Format-CellValue {
    <# A cell value as text for the report: dates in $DateFormat, whole numbers without decimals. #>
    param($Value, [string]$DateFormat = 'yyyy-MM-dd')
    if ($null -eq $Value) { return '' }
    $inv = [Globalization.CultureInfo]::InvariantCulture
    if ($Value -is [DateTime]) {
        if ($Value.TimeOfDay.TotalSeconds -eq 0) { return $Value.ToString($DateFormat, $inv) }
        return $Value.ToString($DateFormat + ' HH:mm', $inv)
    }
    if ($Value -is [double]) {
        if ($Value -eq [math]::Floor($Value) -and [math]::Abs($Value) -lt 1e15) { return ([long]$Value).ToString($inv) }
        return $Value.ToString('0.##########', $inv)
    }
    if ($Value -is [bool]) { if ($Value) { return 'TRUE' } else { return 'FALSE' } }
    ([string]$Value).Trim()
}

function ConvertTo-CaseDate {
    param($Value)
    if ($Value -is [DateTime]) { return $Value.Date }
    if ($Value -is [double]) { return [DateTime]::FromOADate($Value).Date }
    $text = ([string]$Value).Trim()
    $parsed = [DateTime]::MinValue
    $formats = [string[]]@('yyyy-MM-dd', 'dd.MM.yyyy', 'd.M.yyyy', 'dd.MM.yy', 'yyyy-MM-ddTHH:mm:ss', 'dd/MM/yyyy')
    if ([DateTime]::TryParseExact($text, $formats, [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$parsed)) { return $parsed.Date }
    $null
}

# ---------------------------------------------------------------- complaint list

function Get-HeaderRowSetting {
    param([Parameter(Mandatory)]$Config)
    $value = Get-ConfigValue $Config.Settings.Workbook 'HeaderRow'
    if ($null -eq $value -or [string]$value -eq '') { return 1 }
    [int]$value
}

function Import-DossierCases {
    <#
      Reads the complaint list. Maps the columns named in Workbook.Columns to the
      tool's fields and keeps the whole row for the template placeholders.
      Rows without a complaint number or with an unreadable date are left out
      and listed in $script:ImportWarnings. Sorted by date received.
    #>
    param([Parameter(Mandatory)]$Config)
    $script:ImportWarnings = New-Object System.Collections.Generic.List[string]
    $wb = $Config.Settings.Workbook
    $headerRow = Get-HeaderRowSetting $Config
    $rawRows = @(Read-XlsxRows -Path $Config.Paths.Workbook -Sheet $wb.Sheet)
    $header = Get-XlsxHeader $rawRows $headerRow
    if (-not $header.Count) { throw (T 'HeaderRowEmpty' -f $headerRow, $wb.Sheet) }
    $map = [ordered]@{}
    foreach ($field in $script:CaseFields) {
        $column = Get-ConfigString $wb.Columns $field
        if (-not $column) { continue }
        if (-not $header.Contains($column)) { throw (T 'ColumnNotFound' -f $column, "Workbook.Columns.$field", $wb.Sheet, $headerRow, (@($header.Keys) -join ', ')) }
        $map[$field] = $column
    }
    $rows = @(Read-XlsxSheet -Path $Config.Paths.Workbook -Sheet $wb.Sheet -HeaderRow $headerRow)
    $cases = foreach ($row in $rows) {
        $id = Format-CellValue $row.($map.CaseId)
        if (-not $id) { $script:ImportWarnings.Add((T 'RowNoCaseId' -f $row._Row)); continue }
        $date = ConvertTo-CaseDate $row.($map.ReceivedDate)
        if ($null -eq $date) { $script:ImportWarnings.Add((T 'RowBadDate' -f $row._Row, $id, (Format-CellValue $row.($map.ReceivedDate)))); continue }
        $case = [ordered]@{}
        foreach ($field in $script:CaseFields) {
            if ($field -eq 'ReceivedDate') { $case[$field] = $date; continue }
            $value = ''
            if ($map.Contains($field)) { $value = Format-CellValue $row.($map[$field]) }
            $case[$field] = $value
        }
        $case['Row'] = $row
        [pscustomobject]$case
    }
    $cases = @($cases)
    $duplicates = @($cases | Group-Object CaseId | Where-Object Count -gt 1)
    if ($duplicates.Count) { throw (T 'DuplicateIds' -f (($duplicates | ForEach-Object Name) -join ', ')) }
    @($cases | Sort-Object ReceivedDate, CaseId)
}

# ---------------------------------------------------------------- selection

function Get-CaseFolder {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    if (-not $Config.Paths.CaseFolders) { return '' }
    Join-Path $Config.Paths.CaseFolders (Get-SafeFileName $CaseId)
}

function Get-ReportType {
    <#
      Short or long report for one complaint, in this order:
      1. an empty marker file in its case folder (short.txt or long.txt)
      2. the column Workbook.Columns.ReportType (values in Reports.<type>.TypeValues)
      3. Reports.DefaultType
      Returns Type ('Short', 'Long' or $null), Source and Problem.
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Case)
    $reports = $Config.Settings.Reports
    $folder = Get-CaseFolder $Config $Case.CaseId
    if ($folder -and (Test-Path -LiteralPath $folder -PathType Container)) {
        $found = @(foreach ($type in $script:ReportTypes) {
                if (Test-Path -LiteralPath (Join-Path $folder (Get-ConfigString $reports.$type 'Marker')) -PathType Leaf) { $type }
            })
        if ($found.Count -gt 1) {
            return [pscustomobject]@{ Type = $null; Source = ''; Problem = (T 'BothMarkers' -f $reports.Short.Marker, $reports.Long.Marker) }
        }
        if ($found.Count -eq 1) { return [pscustomobject]@{ Type = $found[0]; Source = (T 'SourceMarker' -f $reports.($found[0]).Marker); Problem = '' } }
    }
    $value = [string]$Case.ReportType
    if ($value) {
        foreach ($type in $script:ReportTypes) {
            $accepted = @(@(Get-ConfigList $reports.$type 'TypeValues') + $type | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() })
            if ($accepted -contains $value.ToLowerInvariant()) { return [pscustomobject]@{ Type = $type; Source = (T 'SourceColumn' -f $Config.Settings.Workbook.Columns.ReportType); Problem = '' } }
        }
        return [pscustomobject]@{ Type = $null; Source = ''; Problem = (T 'BadTypeValue' -f $value, $Config.Settings.Workbook.Columns.ReportType) }
    }
    $default = Get-ConfigString $reports 'DefaultType'
    if ($default) { return [pscustomobject]@{ Type = $default; Source = (T 'SourceDefault'); Problem = '' } }
    [pscustomobject]@{ Type = $null; Source = ''; Problem = (T 'NoReportType') }
}

function Get-AuditFolder {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    Join-Path $Config.Paths.Audit (Get-SafeFileName $CaseId)
}

function Get-GeneratedReport {
    <# The audit log of an earlier generated report for this complaint, or $null. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $file = Join-Path (Get-AuditFolder $Config $CaseId) 'audit.json'
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { return $null }
    try { return (Read-JsonFile $file) } catch { return [pscustomobject]@{ generatedAtUtc = '?'; artifacts = $null } }
}

function Get-SelectionPeriod {
    param([Parameter(Mandatory)]$Config, [string]$From, [string]$To)
    $sel = Get-ConfigValue $Config.Settings 'Selection'
    if (-not $From) { $From = Get-ConfigString $sel 'From' }
    if (-not $To) { $To = Get-ConfigString $sel 'To' }
    foreach ($value in @($From, $To)) { if ($value -and -not (Test-DateText $value)) { throw (T 'BadDateSetting' -f 'From/To', $value) } }
    [pscustomobject]@{ From = $From; To = $To }
}

function Select-DossierCases {
    <#
      Decides which complaints get a report in this run.
      Without -CaseId: every complaint received in the period (Selection.From,
      Selection.To, both optional) whose status is in Selection.Statuses (empty
      list = any status). With -CaseId: exactly these complaints.
      A complaint with a generated report (audit log under Output.Path) is
      'done' unless -Regenerate is set. A complaint without a report type is
      'skip'. All others are 'todo'.
    #>
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Cases,
        [string[]]$CaseId,
        [switch]$Regenerate,
        [string]$From,
        [string]$To
    )
    if ($CaseId) {
        $known = @($Cases | ForEach-Object { $_.CaseId })
        $unknown = @($CaseId | Where-Object { $known -notcontains $_ })
        if ($unknown.Count) { throw (T 'UnknownCaseIds' -f ($unknown -join ', ')) }
        $chosen = @($Cases | Where-Object { $CaseId -contains $_.CaseId })
    } else {
        $period = Get-SelectionPeriod $Config $From $To
        $statuses = @(Get-ConfigList (Get-ConfigValue $Config.Settings 'Selection') 'Statuses' | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() })
        $chosen = @($Cases | Where-Object {
                $d = $_.ReceivedDate.ToString('yyyy-MM-dd')
                (-not $period.From -or $d -ge $period.From) -and (-not $period.To -or $d -le $period.To) -and
                (-not $statuses.Count -or $statuses -contains ([string]$_.Status).ToLowerInvariant())
            })
    }
    foreach ($case in $chosen) {
        $existing = Get-GeneratedReport -Config $Config -CaseId $case.CaseId
        if ($existing -and -not $Regenerate) {
            [pscustomobject]@{ Case = $case; Type = $null; State = 'done'; Reason = (T 'AlreadyGenerated' -f $existing.generatedAtUtc) }
            continue
        }
        $type = Get-ReportType -Config $Config -Case $case
        if (-not $type.Type) { [pscustomobject]@{ Case = $case; Type = $null; State = 'skip'; Reason = $type.Problem }; continue }
        [pscustomobject]@{ Case = $case; Type = $type.Type; State = 'todo'; Reason = $type.Source }
    }
}

function Get-SimilarComplaints {
    <#
      Facts the tool computes for one case from the whole complaint list, so the
      model does not have to count:
      - other cases with the same batch (whole list)
      - other cases with the same product group and category, received in the
        LookbackMonths before the case (up to and including its received date),
        and how many of them fall into the last RecentDays
    #>
    param(
        [Parameter(Mandatory)]$Case,
        [Parameter(Mandatory)][object[]]$Cases,
        [int]$LookbackMonths = 12,
        [int]$RecentDays = 30,
        # Ascending thresholds: the level of the highest MinCases that the count reaches.
        [object[]]$FrequencyLevels = @(
            [pscustomobject]@{ MinCases = 0; Level = 'rare' },
            [pscustomobject]@{ MinCases = 1; Level = 'occasional' },
            [pscustomobject]@{ MinCases = 3; Level = 'frequent' },
            [pscustomobject]@{ MinCases = 6; Level = 'very frequent' })
    )
    $others = @($Cases | Where-Object { $_.CaseId -ne $Case.CaseId })
    $batch = ([string](Get-ConfigValue $Case 'Batch')).Trim()
    $sameBatch = @()
    if ($batch) { $sameBatch = @($others | Where-Object { ([string](Get-ConfigValue $_ 'Batch')).Trim() -eq $batch }) }
    $from = $Case.ReceivedDate.AddMonths(-$LookbackMonths)
    $recentFrom = $Case.ReceivedDate.AddDays(-($RecentDays - 1))
    $group = ([string](Get-ConfigValue $Case 'ProductGroup')).Trim()
    $sameKind = @($others | Where-Object {
            ([string](Get-ConfigValue $_ 'ProductGroup')).Trim() -eq $group -and $_.Category -eq $Case.Category -and
            $_.ReceivedDate -gt $from -and $_.ReceivedDate -le $Case.ReceivedDate
        })
    $recent = @($sameKind | Where-Object { $_.ReceivedDate -ge $recentFrom })
    $level = ''
    foreach ($step in @($FrequencyLevels | Sort-Object { [int]$_.MinCases })) { if ($sameKind.Count -ge [int]$step.MinCases) { $level = [string]$step.Level } }
    if (-not $level) { throw 'SimilarComplaints.FrequencyLevels must start at MinCases 0.' }
    $summary = { param($c) [pscustomobject][ordered]@{ CaseId = $c.CaseId; ReceivedDate = $c.ReceivedDate.ToString('yyyy-MM-dd'); Category = $c.Category; ProductGroup = (Get-ConfigValue $c 'ProductGroup'); Batch = (Get-ConfigValue $c 'Batch'); Status = $c.Status } }
    [pscustomobject][ordered]@{
        ListSize                 = $Cases.Count
        Batch                    = $batch
        SameBatch                = @($sameBatch | ForEach-Object { & $summary $_ })
        LookbackMonths           = $LookbackMonths
        LookbackFrom             = $from.AddDays(1).ToString('yyyy-MM-dd')
        LookbackTo               = $Case.ReceivedDate.ToString('yyyy-MM-dd')
        SameProductGroupCategory = @($sameKind | ForEach-Object { & $summary $_ })
        RecentDays               = $RecentDays
        RecentCount              = $recent.Count
        FrequencyLevel           = $level
    }
}

function Get-SimilarComplaintSentences {
    <# The similar-complaint facts as plain sentences, used in the prompt and in the report. #>
    param([Parameter(Mandatory)]$Case, [Parameter(Mandatory)]$Similar)
    $group = [string](Get-ConfigValue $Case 'ProductGroup')
    if (-not $group) { $group = '(no product group)' }
    $count = { param($n) if ($n -eq 0) { 'no other case' } elseif ($n -eq 1) { '1 other case' } else { "$n other cases" } }
    $lines = @()
    $b = @($Similar.SameBatch).Count
    if (-not $Similar.Batch) { $lines += 'Same batch: the case has no batch number.' }
    elseif ($b -eq 0) { $lines += "Same batch: no other case in the complaint list has batch $($Similar.Batch)." }
    else {
        $verb = 'has'
        if ($b -gt 1) { $verb = 'have' }
        $lines += "Same batch: $(& $count $b) in the complaint list $verb batch $($Similar.Batch): $((@($Similar.SameBatch | ForEach-Object { $_.CaseId })) -join ', ')."
    }
    $n = @($Similar.SameProductGroupCategory).Count
    $recent = ''
    if ($n -gt 0 -and $Similar.RecentCount -eq 0) { $recent = "; none of them in the last $($Similar.RecentDays) days" }
    elseif ($n -gt 0) { $recent = "; $($Similar.RecentCount) of them in the last $($Similar.RecentDays) days" }
    $lines += "Same product group and category ($group, $($Case.Category)): $(& $count $n) received from $($Similar.LookbackFrom) to $($Similar.LookbackTo), the $($Similar.LookbackMonths) months up to this case$recent."
    $lines += "Frequency level: $($Similar.FrequencyLevel) (from the number of other cases with the same product group and category in these $($Similar.LookbackMonths) months)."
    $lines += "The complaint list holds $($Similar.ListSize) cases in total."
    $lines
}

# ---------------------------------------------------------------- case documents and uploads

function Protect-PromptText {
    <# Masks e-mail addresses and phone numbers. This is not anonymisation. #>
    param([AllowEmptyString()][string]$Text)
    $Text = [regex]::Replace($Text, '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', '[e-mail removed]')
    [regex]::Replace($Text, '(?<![\w-])(\+|0)\d[\d /-]{6,}\d', '[phone removed]')
}

function Get-MarkerNames {
    param([Parameter(Mandatory)]$Config)
    @($script:ReportTypes | ForEach-Object { (Get-ConfigString $Config.Settings.Reports.$_ 'Marker').ToLowerInvariant() })
}

function Get-CaseDocuments {
    <# Text files from the case folder whose content goes into the prompt. Marker files are left out. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $folder = Get-CaseFolder $Config $CaseId
    if (-not $folder -or -not (Test-Path -LiteralPath $folder -PathType Container)) { return @() }
    $settings = $Config.Settings.CaseFolders
    $extensions = @(Get-ConfigList $settings 'TextExtensions' | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $markers = Get-MarkerNames $Config
    $max = 4000
    if ($null -ne (Get-ConfigValue $settings 'MaxCharsPerDocument')) { $max = [int]$settings.MaxCharsPerDocument }
    foreach ($file in @(Get-ChildItem -LiteralPath $folder -File | Where-Object {
                $extensions -contains $_.Extension.ToLowerInvariant() -and $markers -notcontains $_.Name.ToLowerInvariant() -and $_.Name -notlike '~$*'
            } | Sort-Object Name)) {
        $text = Read-Utf8File $file.FullName
        $truncated = $false
        if ($text.Length -gt $max) { $text = $text.Substring(0, $max); $truncated = $true }
        [pscustomobject]@{ CaseId = $CaseId; Name = $file.Name; Path = $file.FullName; Text = $text.Trim(); Truncated = $truncated }
    }
}

function Get-UploadFiles {
    <#
      Files that the live run attaches to the Copilot chat for one complaint:
      <Copilot.Upload.Path>/<case id>/, or the case folder when Upload.Path is
      empty. Only Copilot.Upload.Extensions; marker files are left out. Fails
      closed on empty files, oversize files or too many files.
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $dir = Get-CaseFolder $Config $CaseId
    if ($Config.Paths.Upload) { $dir = Join-Path $Config.Paths.Upload (Get-SafeFileName $CaseId) }
    if (-not $dir -or -not (Test-Path -LiteralPath $dir -PathType Container)) { return @() }
    $upload = $Config.Settings.Copilot.Upload
    $extensions = @($upload.Extensions | ForEach-Object { $_.ToLowerInvariant() })
    $markers = Get-MarkerNames $Config
    $files = @(Get-ChildItem -LiteralPath $dir -File | Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant() -and $markers -notcontains $_.Name.ToLowerInvariant() -and
            $_.Name -notlike '~$*' -and $_.Name -notlike '.*'
        } | Sort-Object Name)
    if ($files.Count -gt [int]$upload.MaxFiles) { throw "Upload folder of $CaseId holds $($files.Count) files; Copilot.Upload.MaxFiles allows $($upload.MaxFiles)." }
    foreach ($file in $files) {
        if ($file.Length -le 0) { throw "Upload file is empty: $($file.Name)" }
        if ($file.Length -gt ([double]$upload.MaxFileSizeMB * 1MB)) { throw "Upload file is larger than $($upload.MaxFileSizeMB) MB: $($file.Name)" }
    }
    $files
}

# ---------------------------------------------------------------- prompt

function Format-MarkdownTable {
    param([string[]]$Header, [object[]]$Rows)
    $lines = @(
        ('| ' + ($Header -join ' | ') + ' |'),
        ('|' + ((@($Header | ForEach-Object { ' --- ' })) -join '|') + '|')
    )
    foreach ($row in $Rows) { $lines += '| ' + ((@($row | ForEach-Object { ([string]$_).Replace('|', '/') })) -join ' | ') + ' |' }
    $lines -join "`n"
}

function Get-RequiredStatements {
    <#
      Sentences the answer must contain word for word:
      - Reports.<type>.FrequencyStatement with {LEVEL} replaced by the frequency
        level the tool computed (the short report opens its section 2 with it)
      - Reports.<type>.NoSampleStatement when the complaint list says that no
        returned sample was received
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][string]$Type, $Similar)
    $settings = $Config.Settings.Reports.$Type
    $frequency = [string](Get-ConfigValue $settings 'FrequencyStatement')
    if ($frequency -and $null -ne $Similar) { $frequency.Replace('{LEVEL}', [string]$Similar.FrequencyLevel) }
    $statement = [string](Get-ConfigValue $settings 'NoSampleStatement')
    $received = ([string](Get-ConfigValue $Case 'SampleReceived')).Trim()
    if ($statement -and $received -match '^(?i)(no|n|false|0)$') { $statement }
}

function Get-CaseRecordRows {
    param([Parameter(Mandatory)]$Case, [bool]$Mask = $true)
    $description = ($Case.Description -replace '\s+', ' ')
    if ($Mask) { $description = Protect-PromptText $description }
    @(
        , @('Case ID', $Case.CaseId)
        , @('Received', $Case.ReceivedDate.ToString('yyyy-MM-dd'))
        , @('Category', $Case.Category)
        , @('Product', (Get-ConfigValue $Case 'Product'))
        , @('Product group', (Get-ConfigValue $Case 'ProductGroup'))
        , @('Batch', (Get-ConfigValue $Case 'Batch'))
        , @('Severity', (Get-ConfigValue $Case 'Severity'))
        , @('Status', (Get-ConfigValue $Case 'Status'))
        , @('Sample received', (Get-ConfigValue $Case 'SampleReceived'))
        , @('Description', $description)
    )
}

function New-CasePrompt {
    <# Fills the prompt template of the report type with the data of one case. #>
    param(
        [Parameter(Mandatory)][string]$Template,
        [Parameter(Mandatory)]$Case,
        [Parameter(Mandatory)]$Similar,
        [object[]]$Documents = @(),
        [string[]]$Attachments = @(),
        [string[]]$RequiredStatements = @(),
        [string[]]$ProcessFacts = @(),
        [bool]$Mask = $true
    )
    $clean = { param($t) if ($Mask) { Protect-PromptText $t } else { $t } }
    $record = Format-MarkdownTable @('Field', 'Value') (Get-CaseRecordRows $Case $Mask)
    $similarText = (@(Get-SimilarComplaintSentences $Case $Similar) | ForEach-Object { "- $_" }) -join "`n"
    $listed = @(@($Similar.SameBatch) + @($Similar.SameProductGroupCategory) | Sort-Object CaseId -Unique)
    if ($listed.Count) {
        $similarText += "`n`n" + (Format-MarkdownTable @('Case ID', 'Received', 'Category', 'Product group', 'Batch', 'Status') @($listed | ForEach-Object {
                    , @($_.CaseId, $_.ReceivedDate, $_.Category, $_.ProductGroup, $_.Batch, $_.Status)
                }))
    }
    $docText = 'No case documents.'
    if (@($Documents).Count) {
        $docText = (@($Documents | ForEach-Object {
                    $note = ''
                    if ($_.Truncated) { $note = ' (truncated)' }
                    "--- Document $($_.Name)$note ---`n" + (& $clean $_.Text)
                }) -join "`n`n")
    }
    $attachText = 'None.'
    if (@($Attachments).Count) { $attachText = (@($Attachments | ForEach-Object { "- $_" }) -join "`n") }
    $statementText = 'None for this case.'
    if (@($RequiredStatements).Count) { $statementText = (@($RequiredStatements | ForEach-Object { "- $_" }) -join "`n") }
    $factText = 'None.'
    if (@($ProcessFacts).Count) { $factText = (@($ProcessFacts | ForEach-Object { "- $_" }) -join "`n") }

    $values = @{
        '{{CASE_ID}}'             = $Case.CaseId
        '{{REQUIRED_STATEMENTS}}' = $statementText
        '{{PROCESS_FACTS}}'       = $factText
        '{{CASE}}'                = $record
        '{{SIMILAR}}'             = $similarText
        '{{DOCUMENTS}}'           = $docText
        '{{ATTACHMENTS}}'         = $attachText
    }
    $prompt = $Template.Replace("`r`n", "`n")
    foreach ($key in $values.Keys) { $prompt = $prompt.Replace($key, $values[$key]) }
    $left = [regex]::Matches($prompt, '\{\{[A-Z_]+\}\}')
    if ($left.Count) { throw "Prompt template contains unknown placeholder: $($left[0].Value)" }
    $prompt.Trim() + "`n"
}

# ---------------------------------------------------------------- response checks

function ConvertTo-SectionHeadings {
    <#
      Copilot answers arrive as plain text (copy button or UI Automation), without
      Markdown marks; numbered headings may come as list items without their
      number. A line that holds only a section name, with or without its number,
      a list mark, a Markdown heading mark or a trailing colon, becomes "## <section>".
      response.md keeps the answer as received; checks and report use this form.
    #>
    param([AllowEmptyString()][string]$Response, [string[]]$Sections)
    $key = { param($t) ((([string]$t).Trim().TrimEnd(':').Trim() -replace '^(#{1,6}\s+|[-*\u2022]\s+)', '') -replace '^\d+[.)]\s*', '').Trim().ToLowerInvariant() }
    $byKey = @{}
    foreach ($section in $Sections) { $k = & $key $section; if ($k -and -not $byKey.ContainsKey($k)) { $byKey[$k] = $section } }
    $lines = foreach ($line in ($Response -split "`r?`n")) {
        $k = & $key $line
        if ($k -and $byKey.ContainsKey($k)) { "## $($byKey[$k])" } else { $line }
    }
    $lines -join "`n"
}

function Test-ReportResponse {
    <#
      Automated checks on the AI answer for one case. They do not replace the
      human review; the results are shown in the report and the review list and stored in audit.json.
    #>
    param(
        [AllowEmptyString()][string]$Response,
        [string]$Prompt,
        [string]$CaseId,
        [string[]]$RequiredSections,
        [string[]]$KnownCaseIds,
        [string[]]$RequiredStatements = @(),
        # When set, the answer must end with one of these sentences.
        [string[]]$ClassificationStatements = @()
    )
    $checks = New-Object System.Collections.Generic.List[object]
    $checks.Add([pscustomobject]@{ Id = 'not-empty'; Name = 'Response is not empty'; Passed = -not [string]::IsNullOrWhiteSpace($Response); Detail = "$($Response.Length) characters" })
    $echo = (Get-TextSha256 $Response.Trim()) -eq (Get-TextSha256 ([string]$Prompt).Trim())
    $leaked = @($script:PromptHeadings | Where-Object { $Response.Contains($_) })
    $detail = ''
    if ($echo) { $detail = 'response equals the prompt' } elseif ($leaked.Count) { $detail = 'contains prompt heading: ' + ($leaked -join ', ') }
    $checks.Add([pscustomobject]@{ Id = 'not-prompt-copy'; Name = 'Response is not a copy of the prompt'; Passed = (-not $echo -and $leaked.Count -eq 0); Detail = $detail })

    $headings = @([regex]::Matches($Response, '(?m)^#{1,6}\s+(.+?)\s*$') | ForEach-Object { $_.Groups[1].Value.Trim() })
    $missing = @($RequiredSections | Where-Object { $headings -notcontains $_ })
    $detail = 'all present'
    if ($missing.Count) { $detail = 'missing: ' + ($missing -join ', ') }
    $checks.Add([pscustomobject]@{ Id = 'sections'; Name = 'Required sections present'; Passed = ($missing.Count -eq 0); Detail = $detail })

    # A bracketed token with a letter, a digit and a hyphen counts as a case ID
    # citation, e.g. [SYN-2026-004]. Footnotes like [1] and links are ignored.
    $cited = @([regex]::Matches($Response, '\[([A-Za-z][A-Za-z0-9_./]*-[A-Za-z0-9_./-]*\d[A-Za-z0-9_./-]*)\](?!\()') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    # The report must name its own case, cited in brackets or written plainly
    # (for example in the basic data list "Case ID: SYN-2026-007").
    $own = $cited -contains $CaseId
    $detail = "[$CaseId] cited"
    if (-not $own -and $CaseId -and [regex]::IsMatch($Response, '(?<![A-Za-z0-9_-])' + [regex]::Escape($CaseId) + '(?![A-Za-z0-9_-])')) { $own = $true; $detail = "$CaseId named" }
    if (-not $own) { $detail = "$CaseId not named" }
    $checks.Add([pscustomobject]@{ Id = 'own-case-id'; Name = 'Report names its own case ID'; Passed = $own; Detail = $detail })
    $unknown = @($cited | Where-Object { $KnownCaseIds -notcontains $_ })
    $detail = "$($cited.Count) case IDs cited, all in the complaint list"
    if ($unknown.Count) { $detail = 'not in the complaint list: ' + ($unknown -join ', ') }
    $checks.Add([pscustomobject]@{ Id = 'cited-ids'; Name = 'Cited case IDs exist'; Passed = ($unknown.Count -eq 0); Detail = $detail })

    $absent = @($RequiredStatements | Where-Object { $Response.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -lt 0 })
    $detail = 'none required for this case'
    if (@($RequiredStatements).Count) { $detail = 'present: ' + ($RequiredStatements -join ' ') }
    if ($absent.Count) { $detail = 'missing: ' + ($absent -join ' ') }
    $checks.Add([pscustomobject]@{ Id = 'statements'; Name = 'Required statements present'; Passed = ($absent.Count -eq 0); Detail = $detail })

    if (@($ClassificationStatements).Count) {
        $end = $Response.TrimEnd()
        $match = @($ClassificationStatements | Where-Object { $end.EndsWith($_, [StringComparison]::OrdinalIgnoreCase) })
        $detail = 'ends with: ' + ($match -join '')
        if (-not $match.Count) { $detail = 'last sentence is none of: ' + ($ClassificationStatements -join ' / ') }
        $checks.Add([pscustomobject]@{ Id = 'classification'; Name = 'Conclusion ends with a classification'; Passed = ($match.Count -gt 0); Detail = $detail })
    }
    $checks.ToArray()
}


# ---------------------------------------------------------------- OpenXML helpers

function ConvertTo-XmlText {
    param([AllowEmptyString()][string]$Text)
    [Security.SecurityElement]::Escape($Text)
}

function Get-ParagraphText {
    param([string]$ParagraphXml)
    [Net.WebUtility]::HtmlDecode(([regex]::Replace($ParagraphXml, '<[^>]+>', '')))
}

function Update-ZipEntries {
    <# Copies a zip, passing each XML entry's text through $Transform (name, text) -> text. #>
    param([string]$Source, [string]$Destination, [scriptblock]$Transform)
    $tmp = "$Destination.tmp"
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
    $in = [IO.Compression.ZipFile]::OpenRead($Source)
    try {
        $outStream = [IO.File]::Open($tmp, 'CreateNew')
        try {
            $out = New-Object IO.Compression.ZipArchive($outStream, [IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($entry in $in.Entries) {
                    $reader = $entry.Open()
                    $buffer = New-Object IO.MemoryStream
                    try { $reader.CopyTo($buffer) } finally { $reader.Dispose() }
                    $bytes = $buffer.ToArray()
                    if ($entry.FullName -match '\.(xml|rels)$') {
                        $text = $script:Utf8NoBom.GetString($bytes).TrimStart([char]0xFEFF)
                        $text = & $Transform $entry.FullName $text
                        $bytes = $script:Utf8NoBom.GetBytes($text)
                    }
                    $new = $out.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
                    $new.LastWriteTime = $script:FixedZipTime
                    $stream = $new.Open()
                    try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
                }
            } finally { $out.Dispose() }
        } finally { $outStream.Dispose() }
    } catch {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
        throw
    } finally { $in.Dispose() }
    Move-Item -LiteralPath $tmp -Destination $Destination -Force
}

function Get-DocxText {
    <# Visible text of document, headers and footers, one paragraph per line. #>
    param([Parameter(Mandatory)][string]$Path)
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $parts = @($zip.Entries | Where-Object { $_.FullName -match '^word/(document|header\d*|footer\d*)\.xml$' } | Sort-Object FullName)
        $lines = foreach ($entry in $parts) {
            $xml = Get-ZipEntryText $zip $entry.FullName
            foreach ($p in [regex]::Matches($xml, '<w:p(?:\s[^>]*)?>(?:(?!</w:p>).)*?</w:p>', 'Singleline')) { Get-ParagraphText $p.Value }
        }
        $lines -join "`n"
    } finally { $zip.Dispose() }
}

function New-ZipPackage {
    <# Writes an OpenXML package (docx/xlsx) from part name -> XML text, with fixed timestamps. #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][System.Collections.IDictionary]$Parts)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $stream = [IO.File]::Open($Path, 'CreateNew')
    try {
        $zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $Parts.Keys) {
                $entry = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
                $entry.LastWriteTime = $script:FixedZipTime
                $bytes = $script:Utf8NoBom.GetBytes([string]$Parts[$name])
                $s = $entry.Open()
                try { $s.Write($bytes, 0, $bytes.Length) } finally { $s.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $stream.Dispose() }
}

# ---------------------------------------------------------------- report text pieces

function Get-ReportSectionTexts {
    <#
      Splits the answer (after ConvertTo-SectionHeadings, so each section starts
      with "## <heading>") into the text of each configured section. Text before
      the first section is returned under the key '' and is not placed into a
      section placeholder.
    #>
    param([AllowEmptyString()][string]$Response, [string[]]$Headings)
    $buckets = [ordered]@{ '' = New-Object System.Collections.Generic.List[string] }
    foreach ($h in $Headings) { $buckets[$h] = New-Object System.Collections.Generic.List[string] }
    $current = ''
    foreach ($line in ([regex]::Replace($Response, '(?s)<!--.*?-->', '') -split "`r?`n")) {
        if ($line -match '^##\s+(.+?)\s*$' -and $Headings -contains $Matches[1]) { $current = $Matches[1]; continue }
        $buckets[$current].Add($line)
    }
    $result = [ordered]@{}
    foreach ($key in $buckets.Keys) { $result[$key] = (($buckets[$key] -join "`n").Trim()) }
    $result
}

function Get-CheckDisplayName {
    param($Check)
    $key = 'Check_' + $Check.Id
    try { return (T $key) } catch { return $Check.Name }
}

function Get-CheckSummaryText {
    <# The automated check results as lines for the report: a summary line, then one list line per check. #>
    param([Parameter(Mandatory)][object[]]$Checks)
    $failed = @($Checks | Where-Object { -not $_.Passed })
    $lines = @()
    if ($failed.Count) { $lines += (T 'ChecksFailedLine' -f $failed.Count, $Checks.Count) }
    else { $lines += (T 'ChecksPassedLine' -f $Checks.Count) }
    foreach ($c in @($failed) + @($Checks | Where-Object { $_.Passed })) {
        $mark = T 'CheckPassed'
        if (-not $c.Passed) { $mark = T 'CheckFailed' }
        $detail = ''
        if ($c.Detail) { $detail = " ($($c.Detail))" }
        $lines += "- **$mark**: $(Get-CheckDisplayName $c)$detail"
    }
    $lines += (T 'ChecksScope')
    $lines -join "`n"
}

function Get-SimilarComplaintBlock {
    param([Parameter(Mandatory)]$Case, [Parameter(Mandatory)]$Similar)
    $lines = @(Get-SimilarComplaintSentences $Case $Similar | ForEach-Object { "- $_" })
    $listed = @(@($Similar.SameBatch) + @($Similar.SameProductGroupCategory) | Sort-Object CaseId -Unique)
    foreach ($c in $listed) { $lines += "- $($c.CaseId): $($c.ReceivedDate), $($c.Category), $($c.ProductGroup), batch $($c.Batch), $($c.Status)" }
    $lines -join "`n"
}

function Get-TemplateFieldValues {
    <# Template.Fields: placeholder name -> text from the complaint's row (empty cells get Template.EmptyValue). #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Case)
    $template = Get-ConfigValue $Config.Settings 'Template'
    $format = Get-ConfigString $template 'DateFormat'
    if (-not $format) { $format = 'yyyy-MM-dd' }
    $empty = Get-ConfigString $template 'EmptyValue'
    $values = [ordered]@{}
    $fields = Get-ConfigValue $template 'Fields'
    if ($null -eq $fields) { return $values }
    foreach ($p in @($fields.PSObject.Properties)) {
        $column = ([string]$p.Value).Trim()
        $raw = $null
        if ($column -and $Case.Row.PSObject.Properties[$column]) { $raw = $Case.Row.$column }
        $text = Format-CellValue $raw $format
        if (-not $text) { $text = $empty }
        $values[$p.Name] = $text
    }
    $values
}

function Get-ReportFileName {
    param([Parameter(Mandatory)][string]$CaseId, [Parameter(Mandatory)][string]$Type)
    "$(Get-SafeFileName $CaseId)-$($Type.ToLowerInvariant())-report-DRAFT.docx"
}

function Get-DraftNotice {
    param([Parameter(Mandatory)]$Config)
    $text = Get-ConfigString (Get-ConfigValue $Config.Settings 'Review') 'DraftNotice'
    if (-not $text) { $text = T 'DefaultDraftNotice' }
    $text
}

# ---------------------------------------------------------------- one report per complaint

function New-DossierCaseReport {
    <#
      Generates the report for one selected complaint:
      <Output>/reports/<case id>-<type>-report-DRAFT.docx from the user's Word
      template, and <Output>/audit/<case id>/ with case-data.json, prompt.txt,
      response.md and audit.json (SHA-256 of every input and output).
      Demo mode copies <Demo.ResponsePath>/<case id>.md instead of asking Copilot.
    #>
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)]$Selected,
        [Parameter(Mandatory)][object[]]$AllCases,
        [Parameter(Mandatory)][string]$RunId
    )
    $case = $Selected.Case
    $type = $Selected.Type
    $settings = $Config.Settings.Reports.$type
    $root = $Config.Root
    $auditPath = Get-AuditFolder $Config $case.CaseId
    # Leftovers of an interrupted run (no audit.json) are replaced.
    if (Test-Path -LiteralPath $auditPath) { Remove-Item -LiteralPath $auditPath -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $auditPath | Out-Null
    New-Item -ItemType Directory -Force -Path $Config.Paths.Reports | Out-Null

    $events = New-Object System.Collections.Generic.List[object]
    $addEvent = { param($name, $detail) $events.Add([pscustomobject][ordered]@{ atUtc = (Get-UtcStamp); event = $name; detail = $detail }) }
    $inputs = New-Object System.Collections.Generic.List[object]
    $addInput = { param($role, $path) $inputs.Add([pscustomobject][ordered]@{ role = $role; path = (Get-RelativeDisplayPath $root $path); sha256 = (Get-Sha256 $path) }) }
    & $addEvent 'case-started' "$($case.CaseId), $($type.ToLowerInvariant()) report ($($Selected.Reason)), mode=$($Config.Mode.ToLowerInvariant())"
    & $addInput 'config' $Config.Paths.Config
    & $addInput 'complaint-list' $Config.Paths.Workbook
    & $addInput 'prompt-template' $Config.Paths.Prompts.$type
    & $addInput 'word-template' $Config.Paths.Templates.$type

    # 1. Case data: row, similar complaints, documents, uploads
    $lookback = 12
    $recentDays = 30
    $similarSettings = Get-ConfigValue $Config.Settings 'SimilarComplaints'
    if ($null -ne (Get-ConfigValue $similarSettings 'LookbackMonths')) { $lookback = [int]$similarSettings.LookbackMonths }
    if ($null -ne (Get-ConfigValue $similarSettings 'RecentDays')) { $recentDays = [int]$similarSettings.RecentDays }
    $levelArgs = @{}
    if (@(Get-ConfigList $similarSettings 'FrequencyLevels').Count) { $levelArgs.FrequencyLevels = @($similarSettings.FrequencyLevels) }
    $similar = Get-SimilarComplaints -Case $case -Cases $AllCases -LookbackMonths $lookback -RecentDays $recentDays @levelArgs
    $documents = @(Get-CaseDocuments -Config $Config -CaseId $case.CaseId)
    foreach ($doc in $documents) { & $addInput 'case-document' $doc.Path }
    $uploads = @()
    if ($Config.Mode -eq 'Live') {
        $uploads = @(Get-UploadFiles -Config $Config -CaseId $case.CaseId)
        foreach ($file in $uploads) { & $addInput 'upload' $file.FullName }
    }
    $statements = @(Get-RequiredStatements -Config $Config -Case $case -Type $type -Similar $similar)
    $facts = @(Get-ConfigList $settings 'ProcessFacts')
    $fieldValues = Get-TemplateFieldValues -Config $Config -Case $case
    $record = [ordered]@{}
    foreach ($field in $script:CaseFields) {
        if ($field -eq 'ReceivedDate') { $record[$field] = $case.ReceivedDate.ToString('yyyy-MM-dd') } else { $record[$field] = $case.$field }
    }
    $caseData = [ordered]@{
        caseId             = $case.CaseId
        reportType         = $type.ToLowerInvariant()
        reportTypeSource   = $Selected.Reason
        row                = $case.Row._Row
        record             = $record
        templateFields     = $fieldValues
        similarComplaints  = $similar
        requiredStatements = $statements
        processFacts       = $facts
        documents          = @($documents | ForEach-Object { $_.Name })
        uploads            = @($uploads | ForEach-Object { $_.Name })
    }
    Write-JsonFile (Join-Path $auditPath 'case-data.json') $caseData

    # 2. Prompt
    $mask = $true
    if ($null -ne (Get-ConfigValue (Get-ConfigValue $Config.Settings 'Prompt') 'MaskContactData')) { $mask = [bool]$Config.Settings.Prompt.MaskContactData }
    $prompt = New-CasePrompt -Template (Read-Utf8File $Config.Paths.Prompts.$type) -Case $case -Similar $similar -Documents $documents `
        -Attachments @($uploads | ForEach-Object { $_.Name }) -RequiredStatements $statements -ProcessFacts $facts -Mask $mask
    $promptFile = Join-Path $auditPath 'prompt.txt'
    Write-Utf8File $promptFile $prompt
    & $addEvent 'prompt-built' "$($prompt.Length) characters, $($documents.Count) document(s), $($uploads.Count) upload(s)"

    # 3. Response
    $responseFile = Join-Path $auditPath 'response.md'
    $copilotStatus = $null
    if ($Config.Mode -eq 'Demo') {
        $demoFile = Join-Path $Config.Paths.DemoResponses "$(Get-SafeFileName $case.CaseId).md"
        if (-not (Test-Path -LiteralPath $demoFile -PathType Leaf)) { throw (T 'NoDemoAnswer' -f $case.CaseId, (Get-RelativeDisplayPath $root $demoFile -FullPathOutside)) }
        & $addInput 'demo-response' $demoFile
        Write-Utf8File $responseFile (Read-Utf8File $demoFile)
        $modelText = T 'ModelDemo'
        & $addEvent 'response-received' 'demo answer copied'
    } else {
        $cp = $Config.Settings.Copilot
        $modelText = [string]$cp.PreferredModel
        if (Get-ConfigString $cp 'PreferredModelGroup') { $modelText = "$($cp.PreferredModelGroup) / $modelText" }
        $copilotStatus = Invoke-CopilotHandoff -Config $Config -PromptFile $promptFile -ResponseFile $responseFile -UploadFiles $uploads -RunPath $auditPath
        $modelText = "Microsoft 365 Copilot, $modelText"
        & $addEvent 'response-received' "copilot, model=$modelText, uploads=$($uploads.Count)"
    }
    $sections = @(Get-SectionSettings $settings)
    $headings = @($sections | ForEach-Object { $_.Heading })
    # Plain-text answers get their section headings back as "## <heading>".
    $response = ConvertTo-SectionHeadings (Read-Utf8File $responseFile) $headings

    # 4. Checks
    $checks = @(Test-ReportResponse -Response $response -Prompt $prompt -CaseId $case.CaseId -RequiredSections $headings `
            -KnownCaseIds @($AllCases | ForEach-Object { $_.CaseId }) -RequiredStatements $statements `
            -ClassificationStatements @(Get-ConfigList $settings 'ClassificationStatements'))
    # An answer with none of the required sections is not a report (for example
    # a usage-limit message). The complaint then fails without a report, so the
    # next run tries it again, and failures in a row stop the run.
    $present = @($headings | Where-Object { $response -match ('(?m)^##\s+' + [regex]::Escape($_) + '\s*$') })
    if ($headings.Count -and -not $present.Count) {
        $start = (($response -replace '<!--.*?-->', '') -replace '\s+', ' ').Trim()
        if ($start.Length -gt 120) { $start = $start.Substring(0, 120) + '...' }
        throw (T 'NoReportInAnswer' -f $start)
    }
    & $addEvent 'checks-run' "$(@($checks | Where-Object { $_.Passed }).Count) of $($checks.Count) passed"

    # 5. Word report from the user's template
    $template = Get-ConfigValue $Config.Settings 'Template'
    $dateFormat = Get-ConfigString $template 'DateFormat'
    if (-not $dateFormat) { $dateFormat = 'yyyy-MM-dd' }
    $sectionTexts = Get-ReportSectionTexts $response $headings
    $checkText = Get-CheckSummaryText $checks
    $values = @{}
    foreach ($key in $fieldValues.Keys) { $values[$key] = $fieldValues[$key] }
    $typeKey = 'TypeShort'
    if ($type -eq 'Long') { $typeKey = 'TypeLong' }
    $values['DraftNotice'] = Get-DraftNotice $Config
    $values['ReportTitle'] = [string]$settings.Title
    $values['ReportType'] = T $typeKey
    $values['GeneratedAt'] = (Get-Date).ToString($dateFormat + ' HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    $values['Model'] = $modelText
    $values['RunId'] = $RunId
    $values['FrequencyLevel'] = [string]$similar.FrequencyLevel
    $docLines = @(@($documents | ForEach-Object { "- $($_.Name): $(T 'DocInPrompt')" }) + @($uploads | ForEach-Object { "- $($_.Name): $(T 'DocUploaded')" }))
    if (-not $docLines.Count) { $docLines = @(T 'NoDocuments') }
    $blocks = @{
        CheckSummary      = $checkText
        SimilarComplaints = (Get-SimilarComplaintBlock $case $similar)
        CaseDocuments     = ($docLines -join "`n")
        ReportText        = ($response.Trim())
    }
    foreach ($sec in $sections) { if ($sec.Placeholder) { $blocks[$sec.Placeholder] = [string]$sectionTexts[$sec.Heading] } }
    $styles = @{ Heading = 'heading 2'; Bullet = 'List Bullet' }
    $styleSettings = Get-ConfigValue $template 'Styles'
    foreach ($key in 'Heading', 'Bullet') { if (Get-ConfigString $styleSettings $key) { $styles[$key] = Get-ConfigString $styleSettings $key } }
    $status = Get-ConfigString (Get-ConfigValue $Config.Settings 'Review') 'DocumentStatus'
    $reportName = Get-ReportFileName $case.CaseId $type
    $reportFile = Join-Path $Config.Paths.Reports $reportName
    Invoke-WordTemplate -TemplatePath $Config.Paths.Templates.$type -OutputPath $reportFile -Values $values -Blocks $blocks -Styles $styles `
        -Notice $values['DraftNotice'] -Appendix ("## $(T 'ChecksHeading')`n" + $checkText) -ContentStatus $status
    & $addEvent 'report-written' "reports/$reportName"

    # 6. Audit log of this report (SHA-256 hashes). No review data: the review happens in Word.
    # Artifact paths are relative to the output folder.
    $auditRel = 'audit/' + (Get-SafeFileName $case.CaseId)
    $artifact = { param($path, $name) [pscustomobject][ordered]@{ file = $name; sha256 = (Get-Sha256 $path) } }
    $audit = [ordered]@{
        schema         = 'complaint-dossier/audit/v4'
        tool           = [ordered]@{ name = 'Complaint Dossier'; version = (Get-DossierVersion) }
        runId          = $RunId
        caseId         = $case.CaseId
        reportType     = $type.ToLowerInvariant()
        mode           = $Config.Mode.ToLowerInvariant()
        status         = 'DRAFT'
        generatedAtUtc = (Get-UtcStamp)
        summary        = [ordered]@{ receivedDate = $case.ReceivedDate.ToString('yyyy-MM-dd'); product = $case.Product; productGroup = $case.ProductGroup }
        operator       = [ordered]@{ osAccount = [Environment]::UserName; host = 'PowerShell ' + $PSVersionTable.PSVersion.ToString() + ' (' + $PSVersionTable.PSEdition + ')' }
        model          = [ordered]@{ configured = $modelText; copilot = $copilotStatus }
        inputs         = $inputs.ToArray()
        artifacts      = [ordered]@{
            caseData = & $artifact (Join-Path $auditPath 'case-data.json') "$auditRel/case-data.json"
            prompt   = & $artifact $promptFile "$auditRel/prompt.txt"
            response = & $artifact $responseFile "$auditRel/response.md"
            report   = & $artifact $reportFile "reports/$reportName"
        }
        checks         = @($checks | ForEach-Object { [pscustomobject][ordered]@{ id = $_.Id; name = $_.Name; passed = $_.Passed; detail = $_.Detail } })
        review         = [ordered]@{
            where = 'Outside the tool: the reviewer checks the report in Word and finalises it as the QMS requires.'
            note  = 'The report hash above is the generated draft. Edits made during review change the file; that is expected.'
        }
        events         = $events.ToArray()
    }
    $auditFile = Join-Path $auditPath 'audit.json'
    Write-JsonFile $auditFile $audit
    [pscustomobject]@{ CaseId = $case.CaseId; Type = $type; Report = $reportFile; AuditPath = $auditPath; Audit = $auditFile; Checks = $checks }
}

function Move-GeneratedReport {
    <# Before -Regenerate: moves the earlier report and audit folder of a complaint to <Output>/replaced/<time>/. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $audit = Get-AuditFolder $Config $CaseId
    if (-not (Test-Path -LiteralPath $audit)) { return }
    $target = Join-Path (Join-Path $Config.Paths.Output 'replaced') ([DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    foreach ($file in @(Get-ChildItem -LiteralPath $Config.Paths.Reports -File -Filter "$(Get-SafeFileName $CaseId)-*-report-DRAFT.docx" -ErrorAction SilentlyContinue)) {
        Move-Item -LiteralPath $file.FullName -Destination (Join-Path $target $file.Name)
    }
    Move-Item -LiteralPath $audit -Destination (Join-Path $target (Get-SafeFileName $CaseId))
}

# ---------------------------------------------------------------- live Copilot handoff

function Invoke-CopilotHandoff {
    <#
      Runs src/Copilot/Invoke-CopilotRoundtrip.ps1 in Windows PowerShell 5.1
      (STA, .NET Framework UI Automation) for one case. The child opens a new
      chat, selects the configured model (or aborts), uploads the files of the
      case, sends the prompt, waits for the answer and copies it into $ResponseFile.
    #>
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$PromptFile,
        [Parameter(Mandatory)][string]$ResponseFile,
        [Parameter(Mandatory)][string]$RunPath,
        [object[]]$UploadFiles = @(),
        [string]$Executable = 'powershell.exe'
    )
    if ($env:OS -ne 'Windows_NT') { throw 'Live mode needs Windows with the Microsoft 365 Copilot desktop app. Use -Demo on other systems.' }
    $cp = $Config.Settings.Copilot
    $script = Join-Path (Join-Path (Join-Path $Config.Root 'src') 'Copilot') 'Invoke-CopilotRoundtrip.ps1'
    $listFile = Join-Path $RunPath 'uploads.txt'
    Write-Utf8File $listFile ((@($UploadFiles | ForEach-Object { $_.FullName }) -join "`n") + "`n")
    $statusFile = Join-Path $RunPath 'copilot-status.json'

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', $script,
        '-PromptFile', $PromptFile, '-OutputFile', $ResponseFile, '-StatusFile', $statusFile,
        '-AttachmentListFile', $listFile, '-WindowTitle', [string]$cp.WindowTitle, '-PreferredModel', [string]$cp.PreferredModel)
    # Windows PowerShell drops empty arguments, so optional values are only passed when set.
    if (Get-ConfigString $cp 'PreferredModelGroup') { $arguments += @('-PreferredModelGroup', [string]$cp.PreferredModelGroup) }
    if ([bool](Get-ConfigValue $cp 'TemporaryChat')) { $arguments += '-TemporaryChat' }
    foreach ($key in 'StartDelaySeconds', 'NewChatTimeoutSeconds', 'UploadTimeoutSeconds', 'ResponseTimeoutSeconds',
        'StableReadyChecks', 'PollMilliseconds', 'ModelMenuTimeoutMilliseconds') {
        $value = Get-ConfigValue $cp $key
        if ($null -ne $value) { $arguments += @("-$key", [string][int]$value) }
    }
    # The child's progress lines go to a log file, not into the return value.
    $lines = @(& $Executable @arguments 2>&1 | ForEach-Object { [string]$_ })
    $exit = $LASTEXITCODE
    Write-Utf8File (Join-Path $RunPath 'copilot-log.txt') (($lines -join "`n") + "`n")
    $status = $null
    if (Test-Path -LiteralPath $statusFile) { $status = Read-JsonFile $statusFile }
    if ($exit -ne 0) {
        $reason = 'see copilot-log.txt'
        if ($null -ne $status -and $status.error) { $reason = $status.error }
        throw "Copilot handoff failed (exit code $exit): $reason"
    }
    if (-not (Test-Path -LiteralPath $ResponseFile) -or [string]::IsNullOrWhiteSpace((Read-Utf8File $ResponseFile))) {
        throw 'Copilot handoff ended without a response.'
    }
    $status
}

# ---------------------------------------------------------------- review list (Excel)

$script:ReviewColumns = @('CaseId', 'ComplaintDate', 'Product', 'ReportType', 'ReportFile', 'GeneratedAt', 'Mode', 'Checks', 'FailedChecks', 'ReviewedBy', 'ReviewedOn', 'Released', 'Comment')
# Columns the QA team fills in. They are kept when the list is written again.
$script:ReviewKeep = @('ReviewedBy', 'ReviewedOn', 'Released', 'Comment')

function Get-ReviewColumnHeaders {
    <# Column key -> all accepted header texts (English and German), so a list keeps working after a language change. #>
    $null = T 'Col_CaseId'
    $map = @{}
    foreach ($key in $script:ReviewColumns) {
        $map[$key] = @(@('en', 'de') | ForEach-Object { [string]$script:Messages.$_.("Col_$key") } | Where-Object { $_ })
    }
    $map
}

function Read-ReviewEntries {
    <# Reads the reviewer columns of an existing review list: case ID -> @{ key -> value }. #>
    param([Parameter(Mandatory)][string]$Path)
    $kept = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $kept }
    $headers = Get-ReviewColumnHeaders
    $rows = $null
    foreach ($sheet in @('Reports', 'Berichte')) {
        try { $rows = @(Read-XlsxSheet -Path $Path -Sheet $sheet -HeaderRow 1); break } catch { $rows = $null }
    }
    if ($null -eq $rows) { return $kept }
    foreach ($row in $rows) {
        $names = @($row.PSObject.Properties.Name)
        $idColumn = @($names | Where-Object { $headers.CaseId -contains $_ })
        if (-not $idColumn.Count) { continue }
        $id = Format-CellValue $row.($idColumn[0])
        if (-not $id) { continue }
        $entry = @{}
        foreach ($key in $script:ReviewKeep) {
            $column = @($names | Where-Object { $headers[$key] -contains $_ })
            if ($column.Count) { $entry[$key] = $row.($column[0]) }
        }
        $kept[$id] = $entry
    }
    $kept
}

function New-XlsxCell {
    param([int]$Column, [int]$Row, $Value, [int]$Style = 0, [string]$Formula = '')
    $ref = ConvertTo-CellReference $Column $Row
    $s = ''
    if ($Style) { $s = " s=`"$Style`"" }
    if ($Formula) { return "<c r=`"$ref`"$s t=`"str`"><f>$(ConvertTo-XmlText $Formula)</f><v>$(ConvertTo-XmlText ([string]$Value))</v></c>" }
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    if ($Value -is [DateTime]) { return "<c r=`"$ref`" s=`"2`"><v>$($Value.ToOADate().ToString([Globalization.CultureInfo]::InvariantCulture))</v></c>" }
    if ($Value -is [double] -or $Value -is [int]) { return "<c r=`"$ref`"$s><v>$(([double]$Value).ToString('R', [Globalization.CultureInfo]::InvariantCulture))</v></c>" }
    "<c r=`"$ref`"$s t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText ([string]$Value))</t></is></c>"
}

function New-XlsxSheetXml {
    param([string[]]$Header, [object[]]$Rows, [int[]]$Widths, [int]$LinkColumn = -1)
    $ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $cols = for ($i = 0; $i -lt $Widths.Count; $i++) { "<col min=`"$($i + 1)`" max=`"$($i + 1)`" width=`"$($Widths[$i])`" customWidth=`"1`"/>" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<row r="1">' + ((@(for ($i = 0; $i -lt $Header.Count; $i++) { New-XlsxCell $i 1 $Header[$i] 1 })) -join '') + '</row>')
    $r = 2
    foreach ($row in @($Rows | Where-Object { $null -ne $_ })) {
        $cells = for ($i = 0; $i -lt @($row).Count; $i++) {
            if ($i -eq $LinkColumn -and $row[$i]) {
                $name = [string]$row[$i]
                New-XlsxCell $i $r $name 4 ('HYPERLINK("reports\' + $name.Replace('"', '""') + '","' + $name.Replace('"', '""') + '")')
            } else { New-XlsxCell $i $r $row[$i] 3 }
        }
        [void]$sb.Append("<row r=`"$r`">" + ($cells -join '') + '</row>')
        $r++
    }
    $last = ConvertTo-CellReference ($Header.Count - 1) ([math]::Max(1, $r - 1))
    "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<worksheet $ns><sheetViews><sheetView workbookViewId=`"0`"><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews><cols>$($cols -join '')</cols><sheetData>$($sb.ToString())</sheetData><autoFilter ref=`"A1:$last`"/></worksheet>"
}

function Update-ReviewList {
    <#
      Writes <Output>/review-list.xlsx: one row per generated report (from the
      audit logs), with the automated check results and empty columns
      "Reviewed by", "Reviewed on", "Released", "Comment" for the QA team.
      What the team entered there is kept when the list is written again.
      A second sheet lists the complaints of the last run that got no report.
      When the list is open in Excel, a copy with a time stamp is written instead.
    #>
    param([Parameter(Mandatory)]$Config, [object[]]$NotGenerated = @())
    $path = $Config.Paths.ReviewList
    $kept = Read-ReviewEntries $path
    $entries = foreach ($file in @(Get-ChildItem -LiteralPath $Config.Paths.Audit -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'audit.json' } | Where-Object { Test-Path -LiteralPath $_ })) {
        try { Read-JsonFile $file } catch { }
    }
    $rows = @(foreach ($a in @($entries | Where-Object { $_ } | Sort-Object { $_.summary.receivedDate }, caseId)) {
        $checks = @($a.checks)
        $failed = @($checks | Where-Object { -not $_.passed })
        $typeKey = 'TypeShort'
        if ($a.reportType -eq 'long') { $typeKey = 'TypeLong' }
        $mode = $a.model.configured
        $k = $kept[[string]$a.caseId]
        if ($null -eq $k) { $k = @{} }
        , @(
            [string]$a.caseId,
            [string]$a.summary.receivedDate,
            [string]$a.summary.product,
            (T $typeKey),
            ([string]$a.artifacts.report.file -replace '^reports/', ''),
            ([string]$a.generatedAtUtc).Replace('T', ' ').Replace('Z', ' UTC'),
            [string]$mode,
            (T 'ChecksCell' -f ($checks.Count - $failed.Count), $checks.Count),
            ((@($failed | ForEach-Object { $_.name + $(if ($_.detail) { " ($($_.detail))" } else { '' }) })) -join '; '),
            $k['ReviewedBy'], $k['ReviewedOn'], $k['Released'], $k['Comment']
        )
    })
    $header = @($script:ReviewColumns | ForEach-Object { T "Col_$_" })
    $missing = @(foreach ($n in @($NotGenerated | Where-Object { $_ })) { , @([string]$n.CaseId, [string]$n.ReceivedDate, [string]$n.State, [string]$n.Reason) })
    $parts = [ordered]@{}
    $parts['[Content_Types].xml'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@
    $parts['_rels/.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
    $s1 = ConvertTo-XmlText (T 'SheetReports')
    $s2 = ConvertTo-XmlText (T 'SheetNotGenerated')
    $parts['xl/workbook.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets><sheet name=`"$s1`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"$s2`" sheetId=`"2`" r:id=`"rId2`"/></sheets></workbook>"
    $parts['xl/_rels/workbook.xml.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@
    $parts['xl/styles.xml'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>
</styleSheet>
'@
    $parts['xl/worksheets/sheet1.xml'] = New-XlsxSheetXml -Header $header -Rows $rows -Widths @(15, 13, 24, 13, 34, 19, 26, 12, 40, 18, 13, 12, 30) -LinkColumn 4
    $parts['xl/worksheets/sheet2.xml'] = New-XlsxSheetXml -Header @((T 'Col_CaseId'), (T 'Col_ComplaintDate'), (T 'Col_State'), (T 'Col_Reason')) -Rows $missing -Widths @(15, 13, 14, 90)
    New-Item -ItemType Directory -Force -Path $Config.Paths.Output | Out-Null
    try {
        $tmp = "$path.new"
        New-ZipPackage -Path $tmp -Parts $parts
        Move-Item -LiteralPath $tmp -Destination $path -Force
        return $path
    } catch {
        $alt = Join-Path $Config.Paths.Output ('review-list-' + (Get-Date).ToString('yyyyMMdd-HHmmss') + '.xlsx')
        if (Test-Path -LiteralPath $tmp) { Move-Item -LiteralPath $tmp -Destination $alt -Force } else { New-ZipPackage -Path $alt -Parts $parts }
        Write-Warning (T 'ReviewListLocked' -f $path, $alt)
        return $alt
    }
}

# ---------------------------------------------------------------- setup checks

function Get-EditDistance {
    param([string]$A, [string]$B)
    $A = $A.ToLowerInvariant(); $B = $B.ToLowerInvariant()
    $prev = 0..$B.Length
    for ($i = 1; $i -le $A.Length; $i++) {
        $cur = @($i) + @(0) * $B.Length
        for ($j = 1; $j -le $B.Length; $j++) {
            $cost = 1
            if ($A[$i - 1] -eq $B[$j - 1]) { $cost = 0 }
            $cur[$j] = [math]::Min([math]::Min($cur[$j - 1] + 1, $prev[$j] + 1), $prev[$j - 1] + $cost)
        }
        $prev = $cur
    }
    $prev[$B.Length]
}

function Get-ClosestName {
    param([string]$Name, [string[]]$Candidates)
    $best = ''; $bestDistance = 4
    foreach ($c in $Candidates) {
        $d = Get-EditDistance $Name $c
        if ($d -lt $bestDistance) { $best = $c; $bestDistance = $d }
    }
    $best
}

function Get-KnownPlaceholders {
    <# Placeholder names a template of this report type may use: name -> 'inline' or 'block'. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Type)
    $known = @{}
    foreach ($key in $script:ToolPlaceholders.Keys) { $known[$key] = $script:ToolPlaceholders[$key] }
    $fields = Get-ConfigValue (Get-ConfigValue $Config.Settings 'Template') 'Fields'
    if ($fields) { foreach ($p in @($fields.PSObject.Properties)) { $known[$p.Name] = 'inline' } }
    foreach ($sec in @(Get-SectionSettings $Config.Settings.Reports.$Type)) { if ($sec.Placeholder) { $known[$sec.Placeholder] = 'block' } }
    $known
}

function Test-CopilotWindowOpen {
    param([string]$Title)
    $escaped = [regex]::Escape($Title)
    @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -match ('(?i)^' + $escaped + '(?:\s*[-|].*)?$') }).Count -gt 0
}

function Test-DossierSetup {
    <#
      Everything Setup checks, as findings (Level ok/info/warn/error, Section,
      Text, Fix): configuration, Excel list (sheet, header row, columns, rows),
      Word templates (placeholders), Copilot (live mode) and the output folder.
      The run refuses to start while a finding has the level 'error'.
    #>
    param([Parameter(Mandatory)]$Config, [string]$From, [string]$To)
    $c = $Config.Settings
    $p = $Config.Paths
    $configFindings = @(Test-DossierConfig $Config)
    $configFindings

    # Excel list
    $wb = Get-ConfigValue $c 'Workbook'
    $excelBlocked = @($configFindings | Where-Object { $_.Level -eq 'error' -and $_.Section -eq 'excel' }).Count -gt 0
    if (-not $excelBlocked) {
        $sheet = Get-ConfigString $wb 'Sheet'
        $headerRow = Get-HeaderRowSetting $Config
        $rows = $null
        try { $rows = @(Read-XlsxRows -Path $p.Workbook -Sheet $sheet) }
        catch {
            $names = @()
            try { $names = @(Get-XlsxSheetNames $p.Workbook) } catch { }
            New-Finding error 'excel' $_.Exception.Message (T 'FixSheet' -f ($names -join ', '))
        }
        if ($null -ne $rows) {
            $header = Get-XlsxHeader $rows $headerRow
            $headerNames = @($header.Keys)
            if (-not $header.Count) {
                New-Finding error 'excel' (T 'HeaderRowEmpty' -f $headerRow, $sheet) (T 'FixHeaderRow' -f (Find-XlsxHeaderRow $rows))
            } else {
                New-Finding ok 'excel' (T 'ExcelOk' -f (Get-RelativeDisplayPath $Config.Root $p.Workbook -FullPathOutside), $sheet, $headerRow, $header.Count) ''
                $columnErrors = 0
                foreach ($field in $script:CaseFields) {
                    $column = Get-ConfigString $wb.Columns $field
                    if (-not $column) { continue }
                    if (-not $header.Contains($column)) {
                        $columnErrors++
                        $guess = Get-ClosestName $column $headerNames
                        $hint = ''
                        if ($guess) { $hint = ' ' + (T 'DidYouMean' -f $guess) }
                        New-Finding error 'excel' ((T 'ColumnMissing' -f $column, (T ('Role' + $field))) + $hint) (T 'FixColumn' -f "Workbook.Columns.$field", ($headerNames -join ' | '))
                    }
                }
                $fields = Get-ConfigValue (Get-ConfigValue $c 'Template') 'Fields'
                if ($fields) {
                    foreach ($f in @($fields.PSObject.Properties)) {
                        $column = ([string]$f.Value).Trim()
                        if (-not $column -or $header.Contains($column)) { continue }
                        $columnErrors++
                        $guess = Get-ClosestName $column $headerNames
                        $hint = ''
                        if ($guess) { $hint = ' ' + (T 'DidYouMean' -f $guess) }
                        New-Finding error 'excel' ((T 'FieldColumnMissing' -f $column, $f.Name) + $hint) (T 'FixColumn' -f "Template.Fields.$($f.Name)", ($headerNames -join ' | '))
                    }
                }
                if (-not $columnErrors) {
                    New-Finding ok 'excel' (T 'ColumnsOk') ''
                    try {
                        $cases = @(Import-DossierCases -Config $Config)
                        $warnings = @($script:ImportWarnings)
                        New-Finding ok 'excel' (T 'CasesRead' -f $cases.Count) ''
                        foreach ($w in @($warnings | Select-Object -First 10)) { New-Finding warn 'excel' $w (T 'FixRowSkipped') }
                        if ($warnings.Count -gt 10) { New-Finding warn 'excel' (T 'MoreRowsSkipped' -f ($warnings.Count - 10)) '' }
                        $period = Get-SelectionPeriod $Config $From $To
                        $selection = @(Select-DossierCases -Config $Config -Cases $cases -From $period.From -To $period.To)
                        $todo = @($selection | Where-Object { $_.State -eq 'todo' })
                        $done = @($selection | Where-Object { $_.State -eq 'done' })
                        $skip = @($selection | Where-Object { $_.State -eq 'skip' })
                        $fromText = $period.From; if (-not $fromText) { $fromText = '...' }
                        $toText = $period.To; if (-not $toText) { $toText = '...' }
                        $level = 'ok'
                        if (-not $selection.Count) { $level = 'warn' }
                        New-Finding $level 'excel' (T 'PeriodCount' -f $fromText, $toText, $selection.Count, $todo.Count, $done.Count, $skip.Count) ''
                        foreach ($group in @($skip | Group-Object Reason)) {
                            New-Finding warn 'excel' (T 'SkipGroup' -f $group.Count, $group.Name, ((@($group.Group | Select-Object -First 5 | ForEach-Object { $_.Case.CaseId })) -join ', ')) (T 'FixReportType')
                        }
                    } catch { New-Finding error 'excel' $_.Exception.Message '' }
                }
            }
        }
    }

    # Word templates
    $checked = @{}
    foreach ($type in $script:ReportTypes) {
        $template = $p.Templates.$type
        if (-not $template -or -not (Test-Path -LiteralPath $template -PathType Leaf)) { continue }
        if ([IO.Path]::GetExtension($template).ToLowerInvariant() -notin '.docx', '.dotx') { continue }
        $display = Get-RelativeDisplayPath $Config.Root $template -FullPathOutside
        $typeText = T ('Type' + $type)
        try { $found = @(Get-WordPlaceholders $template) }
        catch { New-Finding error 'templates' (T 'TemplateUnreadable' -f $display, $_.Exception.Message) (T 'FixSaveAsDocx'); continue }
        $known = Get-KnownPlaceholders $Config $type
        $names = @($found | ForEach-Object { $_.Name } | Sort-Object -Unique)
        $errors = 0
        foreach ($name in $names) {
            if ($known.ContainsKey($name)) { continue }
            $errors++
            $guess = Get-ClosestName $name @($known.Keys)
            $hint = ''
            if ($guess) { $hint = ' ' + (T 'DidYouMean' -f "{{$guess}}") }
            New-Finding error 'templates' ((T 'PlaceholderUnknown' -f "{{$name}}", $display) + $hint) (T 'FixPlaceholderUnknown')
        }
        $hasReportText = @($names | Where-Object { $_ -eq 'ReportText' }).Count -gt 0
        foreach ($sec in @(Get-SectionSettings $c.Reports.$type)) {
            $placed = $sec.Placeholder -and @($names | Where-Object { $_ -eq $sec.Placeholder }).Count -gt 0
            if (-not $placed -and -not $hasReportText) {
                $errors++
                $ph = $sec.Placeholder
                if (-not $ph) { $ph = '?' }
                New-Finding error 'templates' (T 'SectionNoPlace' -f $sec.Heading, $display, "{{$ph}}") (T 'FixSectionNoPlace' -f "{{$ph}}")
            }
        }
        foreach ($occ in @($found | Where-Object { $known.ContainsKey($_.Name) -and $known[$_.Name] -eq 'block' -and -not $_.Alone })) {
            New-Finding warn 'templates' (T 'BlockInline' -f "{{$($occ.Name)}}", $display) (T 'FixBlockInline')
        }
        if (@($names | Where-Object { $_ -eq 'DraftNotice' }).Count -eq 0) { New-Finding warn 'templates' (T 'NoDraftNotice' -f $display) (T 'FixDraftNotice') }
        if (@($names | Where-Object { $_ -eq 'CheckSummary' }).Count -eq 0) { New-Finding warn 'templates' (T 'NoCheckSummary' -f $display) (T 'FixCheckSummary') }
        if (-not $errors) { New-Finding ok 'templates' (T 'TemplateOk' -f $typeText, $display, $names.Count) '' }
        $checked[$template] = $true
    }

    # Output folder
    try {
        New-Item -ItemType Directory -Force -Path $p.Output -ErrorAction Stop | Out-Null
        New-Finding ok 'output' (T 'OutputOk' -f (Get-RelativeDisplayPath $Config.Root $p.Output -FullPathOutside)) ''
    } catch { New-Finding error 'output' (T 'OutputNotWritable' -f $p.Output, $_.Exception.Message) '' }
}

Export-ModuleMember -Function *-*
