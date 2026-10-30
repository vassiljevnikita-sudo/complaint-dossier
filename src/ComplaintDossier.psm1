# Complaint Dossier core functions.
# Runs on PowerShell 7 (Windows, Linux, macOS) and on Windows PowerShell 5.1.
# Keep this file ASCII-only: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.

Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:ToolVersion = '0.2.0'
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:FixedZipTime = [DateTimeOffset]::new(2026, 1, 1, 0, 0, 0, [TimeSpan]::Zero)

# Text markers that Approve-ComplaintDossier replaces. Each must appear in the draft.
$script:Markers = [ordered]@{
    Banner    = 'DRAFT - NOT REVIEWED - NOT FOR USE'
    Status    = 'Status: DRAFT (awaiting human review)'
    Reviewer  = 'Reviewer: (not signed)'
    SignedAt  = 'Signed at (UTC): (not signed)'
    Comment   = 'Review comment: (none)'
    Exception = 'Accepted check failures: (not signed)'
}

function Get-DossierVersion { $script:ToolVersion }
function Get-DossierMarkers { $script:Markers }

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

# ---------------------------------------------------------------- configuration

function Get-ConfigValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    $property.Value
}

function Import-DossierConfig {
    <#
      Loads and validates the JSON configuration. Every path the tool touches
      comes from this file. Fails closed: a missing required value stops the run.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root,
        [ValidateSet('Demo', 'Live')][string]$Mode = 'Demo'
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Configuration file not found: $Path. Copy config.example.json to config.json and edit it."
    }
    try { $config = Read-Utf8File $Path | ConvertFrom-Json }
    catch { throw "Configuration file is not valid JSON: $Path. $($_.Exception.Message)" }

    $required = @(
        @('Workbook', 'Path'), @('Workbook', 'Sheet'), @('Workbook', 'Columns'),
        @('Prompt', 'TemplatePath'), @('Report', 'TemplatePath'), @('Report', 'OutputPath')
    )
    if ($Mode -eq 'Demo') { $required += , @('Demo', 'ResponsePath') }
    if ($Mode -eq 'Live') { $required += @(@('Copilot', 'WindowTitle'), @('Copilot', 'PreferredModel')) }
    foreach ($pair in $required) {
        $value = Get-ConfigValue (Get-ConfigValue $config $pair[0]) $pair[1]
        if ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) {
            $hint = ''
            if ($pair[1] -eq 'PreferredModel') { $hint = ' Enter the exact model name shown in the Copilot model menu. The tool never picks a model on its own.' }
            throw "Configuration value missing: $($pair[0]).$($pair[1]).$hint"
        }
    }
    foreach ($field in 'CaseId', 'ReceivedDate', 'Category', 'Description') {
        if ([string]::IsNullOrWhiteSpace([string](Get-ConfigValue $config.Workbook.Columns $field))) {
            throw "Configuration value missing: Workbook.Columns.$field"
        }
    }

    $paths = [ordered]@{
        Config         = [IO.Path]::GetFullPath($Path)
        Workbook       = Resolve-DossierPath $Root $config.Workbook.Path
        Documents      = Resolve-DossierPath $Root ([string](Get-ConfigValue $config.Documents 'Path'))
        PromptTemplate = Resolve-DossierPath $Root $config.Prompt.TemplatePath
        ReportTemplate = Resolve-DossierPath $Root $config.Report.TemplatePath
        Output         = Resolve-DossierPath $Root $config.Report.OutputPath
        DemoResponse   = Resolve-DossierPath $Root ([string](Get-ConfigValue $config.Demo 'ResponsePath'))
        Upload         = Resolve-DossierPath $Root ([string](Get-ConfigValue (Get-ConfigValue $config.Copilot 'Upload') 'Path'))
    }
    foreach ($key in 'Workbook', 'PromptTemplate', 'ReportTemplate') {
        if (-not (Test-Path -LiteralPath $paths[$key] -PathType Leaf)) { throw "Configured file not found ($key): $($paths[$key])" }
    }
    if ($Mode -eq 'Demo' -and -not (Test-Path -LiteralPath $paths.DemoResponse -PathType Leaf)) {
        throw "Configured file not found (Demo.ResponsePath): $($paths.DemoResponse)"
    }
    [pscustomobject]@{ Settings = $config; Paths = [pscustomobject]$paths; Root = $Root; Mode = $Mode }
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

function Read-XlsxSheet {
    <#
      Reads one worksheet into objects keyed by the header row.
      Pure OpenXML: works without Excel, on Windows and Linux.
      Numbers are returned as [double], text as [string].
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Sheet,
        [int]$HeaderRow = 1
    )
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        $zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
        $workbook = [xml](Get-ZipEntryText $zip 'xl/workbook.xml')
        $sheetNode = @($workbook.SelectNodes("//*[local-name()='sheet']") | Where-Object { $_.GetAttribute('name') -eq $Sheet })
        if ($sheetNode.Count -ne 1) {
            $names = @($workbook.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') }) -join ', '
            throw "Worksheet '$Sheet' not found in $([IO.Path]::GetFileName($Path)). Available: $names"
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

        $sheetXml = [xml](Get-ZipEntryText $zip $target)
        $rows = New-Object System.Collections.Generic.List[object]
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
                    default { if ($v -and $v.InnerText -ne '') { $value = [double]::Parse($v.InnerText, [Globalization.CultureInfo]::InvariantCulture) } }
                }
                $values[$col] = $value
            }
            $rows.Add([pscustomobject]@{ Number = [int]$row.GetAttribute('r'); Values = $values })
        }
    } finally { $stream.Dispose() }

    $header = @($rows | Where-Object { $_.Number -eq $HeaderRow })
    if ($header.Count -ne 1) { throw "Header row $HeaderRow not found in worksheet '$Sheet'." }
    $columns = @{}
    foreach ($key in $header[0].Values.Keys) {
        $name = ([string]$header[0].Values[$key]).Trim()
        if ($name) { $columns[$name] = $key }
    }
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

function ConvertTo-CaseDate {
    param($Value, [string]$CaseId)
    if ($Value -is [double]) { return [DateTime]::FromOADate($Value).Date }
    $text = [string]$Value
    $parsed = [DateTime]::MinValue
    $formats = [string[]]@('yyyy-MM-dd', 'dd.MM.yyyy', 'yyyy-MM-ddTHH:mm:ss')
    if ([DateTime]::TryParseExact($text.Trim(), $formats, [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$parsed)) { return $parsed.Date }
    throw "Case $CaseId has an unreadable ReceivedDate: '$text' (use an Excel date or yyyy-MM-dd)."
}

function Import-DossierCases {
    <# Reads the case table and maps workbook columns to the neutral case schema. #>
    param([Parameter(Mandatory)]$Config, [string[]]$CaseId)
    $wb = $Config.Settings.Workbook
    $headerRow = 1
    if ($null -ne (Get-ConfigValue $wb 'HeaderRow')) { $headerRow = [int]$wb.HeaderRow }
    $rows = @(Read-XlsxSheet -Path $Config.Paths.Workbook -Sheet $wb.Sheet -HeaderRow $headerRow)
    if ($rows.Count -eq 0) { throw "Worksheet '$($wb.Sheet)' contains no cases." }

    $map = [ordered]@{}
    foreach ($field in 'CaseId', 'ReceivedDate', 'Category', 'ProductGroup', 'Severity', 'Status', 'Description') {
        $column = [string](Get-ConfigValue $wb.Columns $field)
        if ([string]::IsNullOrWhiteSpace($column)) { continue }
        if (-not ($rows[0].PSObject.Properties.Name -contains $column)) {
            throw "Column '$column' (mapped to $field) not found in worksheet '$($wb.Sheet)', header row $headerRow."
        }
        $map[$field] = $column
    }

    $cases = foreach ($row in $rows) {
        $id = ([string]$row.($map.CaseId)).Trim()
        if (-not $id) { throw "Row $($row._Row) has no case ID." }
        $case = [ordered]@{}
        foreach ($field in 'CaseId', 'ReceivedDate', 'Category', 'ProductGroup', 'Severity', 'Status', 'Description') {
            $value = ''
            if ($map.Contains($field)) { $value = $row.($map[$field]) }
            if ($field -eq 'ReceivedDate') { $value = ConvertTo-CaseDate $value $id }
            elseif ($null -eq $value) { $value = '' }
            else { $value = ([string]$value).Trim() }
            $case[$field] = $value
        }
        if (-not $case.Category) { throw "Case $id has no category." }
        [pscustomobject]$case
    }
    $duplicates = @($cases | Group-Object CaseId | Where-Object Count -gt 1)
    if ($duplicates.Count) { throw "Duplicate case IDs in workbook: $(($duplicates | ForEach-Object Name) -join ', ')" }

    if ($CaseId) {
        $unknown = @($CaseId | Where-Object { @($cases.CaseId) -notcontains $_ })
        if ($unknown.Count) { throw "Case ID not found in workbook: $($unknown -join ', ')" }
        $cases = @($cases | Where-Object { $CaseId -contains $_.CaseId })
    }
    @($cases | Sort-Object ReceivedDate, CaseId)
}

# ---------------------------------------------------------------- statistics

function Get-DossierStatistics {
    <# Counts by category and a month-by-month trend (months without cases count as 0). #>
    param([Parameter(Mandatory)][object[]]$Cases)
    $total = $Cases.Count
    $byCategory = @($Cases | Group-Object Category | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name | ForEach-Object {
            [pscustomobject]@{ Category = $_.Name; Count = $_.Count; Share = [math]::Round(100.0 * $_.Count / $total, 1) }
        })
    $dates = @($Cases | ForEach-Object { $_.ReceivedDate } | Sort-Object)
    $first = $dates[0]
    $last = $dates[-1]
    $month = New-Object DateTime($first.Year, $first.Month, 1)
    $end = New-Object DateTime($last.Year, $last.Month, 1)
    $byMonth = while ($month -le $end) {
        $key = $month.ToString('yyyy-MM')
        [pscustomobject]@{ Month = $key; Count = @($Cases | Where-Object { $_.ReceivedDate.ToString('yyyy-MM') -eq $key }).Count }
        $month = $month.AddMonths(1)
    }
    [pscustomobject]@{
        Total      = $total
        From       = $first.ToString('yyyy-MM-dd')
        To         = $last.ToString('yyyy-MM-dd')
        ByCategory = @($byCategory)
        ByMonth    = @($byMonth)
    }
}

# ---------------------------------------------------------------- documents and uploads

function Protect-PromptText {
    <# Masks e-mail addresses and phone numbers. This is not anonymisation. #>
    param([AllowEmptyString()][string]$Text)
    $Text = [regex]::Replace($Text, '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', '[e-mail removed]')
    [regex]::Replace($Text, '(?<![\w-])(\+|0)\d[\d /-]{6,}\d', '[phone removed]')
}

function Get-CaseDocuments {
    <# Text documents from <Documents.Path>/<CaseId>/ that go into the prompt. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][object[]]$Cases)
    $root = $Config.Paths.Documents
    if (-not $root) { return @() }
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "Configured documents folder not found: $root" }
    $extensions = @($Config.Settings.Documents.Extensions | ForEach-Object { $_.ToLowerInvariant() })
    $max = 4000
    if ($null -ne (Get-ConfigValue $Config.Settings.Documents 'MaxCharsPerDocument')) { $max = [int]$Config.Settings.Documents.MaxCharsPerDocument }
    foreach ($case in $Cases) {
        $folder = Join-Path $root $case.CaseId
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -File | Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } | Sort-Object Name)) {
            $text = Read-Utf8File $file.FullName
            $truncated = $false
            if ($text.Length -gt $max) { $text = $text.Substring(0, $max); $truncated = $true }
            [pscustomobject]@{ CaseId = $case.CaseId; Name = $file.Name; Path = $file.FullName; Text = $text.Trim(); Truncated = $truncated }
        }
    }
}

function Get-UploadFiles {
    <#
      Files from Copilot.Upload.Path that the live run attaches to the chat.
      Fails closed on empty files, oversize files or too many files.
    #>
    param([Parameter(Mandatory)]$Config)
    $dir = $Config.Paths.Upload
    if (-not $dir) { return @() }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { throw "Configured upload folder not found: $dir" }
    $upload = $Config.Settings.Copilot.Upload
    $extensions = @($upload.Extensions | ForEach-Object { $_.ToLowerInvariant() })
    $files = @(Get-ChildItem -LiteralPath $dir -File | Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant() -and $_.Name -notlike '~$*' -and $_.Name -notlike '.*'
        } | Sort-Object Name)
    if ($files.Count -gt [int]$upload.MaxFiles) { throw "Upload folder holds $($files.Count) files; Copilot.Upload.MaxFiles allows $($upload.MaxFiles)." }
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

function New-DossierPrompt {
    param(
        [Parameter(Mandatory)][string]$Template,
        [Parameter(Mandatory)][object[]]$Cases,
        [Parameter(Mandatory)]$Statistics,
        [object[]]$Documents = @(),
        [string[]]$Attachments = @(),
        [bool]$Mask = $true
    )
    $clean = { param($t) if ($Mask) { Protect-PromptText $t } else { $t } }
    $stats = "Cases by category:`n" + (Format-MarkdownTable @('Category', 'Count', 'Share %') @($Statistics.ByCategory | ForEach-Object { , @($_.Category, $_.Count, $_.Share.ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)) }))
    $stats += "`n`nCases by month:`n" + (Format-MarkdownTable @('Month', 'Count') @($Statistics.ByMonth | ForEach-Object { , @($_.Month, $_.Count) }))
    $caseTable = Format-MarkdownTable @('Case ID', 'Received', 'Category', 'Product group', 'Severity', 'Status', 'Description') @($Cases | ForEach-Object {
            , @($_.CaseId, $_.ReceivedDate.ToString('yyyy-MM-dd'), $_.Category, $_.ProductGroup, $_.Severity, $_.Status, (& $clean (($_.Description -replace '\s+', ' '))))
        })
    $docText = 'No case documents.'
    if (@($Documents).Count) {
        $docText = (@($Documents | ForEach-Object {
                    $note = ''
                    if ($_.Truncated) { $note = ' (truncated)' }
                    "--- Document $($_.Name) for case $($_.CaseId)$note ---`n" + (& $clean $_.Text)
                }) -join "`n`n")
    }
    $attachText = 'None.'
    if (@($Attachments).Count) { $attachText = (@($Attachments | ForEach-Object { "- $_" }) -join "`n") }

    $values = @{
        '{{PERIOD}}'      = "$($Statistics.From) to $($Statistics.To)"
        '{{CASE_COUNT}}'  = [string]$Statistics.Total
        '{{STATISTICS}}'  = $stats
        '{{CASES}}'       = $caseTable
        '{{DOCUMENTS}}'   = $docText
        '{{ATTACHMENTS}}' = $attachText
    }
    $prompt = $Template.Replace("`r`n", "`n")
    foreach ($key in $values.Keys) { $prompt = $prompt.Replace($key, $values[$key]) }
    $left = [regex]::Matches($prompt, '\{\{[A-Z_]+\}\}')
    if ($left.Count) { throw "Prompt template contains unknown placeholder: $($left[0].Value)" }
    $prompt.Trim() + "`n"
}

# ---------------------------------------------------------------- response checks

function Test-DossierResponse {
    <#
      Automated checks on the AI response. They do not replace the human review;
      the results are printed in the report and stored in the audit log.
    #>
    param([string]$Response, [string]$Prompt, [string[]]$RequiredSections, [string[]]$KnownCaseIds)
    $checks = New-Object System.Collections.Generic.List[object]
    $checks.Add([pscustomobject]@{ Name = 'Response is not empty'; Passed = -not [string]::IsNullOrWhiteSpace($Response); Detail = "$($Response.Length) characters" })
    $echo = (Get-TextSha256 $Response.Trim()) -eq (Get-TextSha256 $Prompt.Trim())
    $checks.Add([pscustomobject]@{ Name = 'Response differs from prompt'; Passed = -not $echo; Detail = '' })

    $headings = @([regex]::Matches($Response, '(?m)^#{1,6}\s+(.+?)\s*$') | ForEach-Object { $_.Groups[1].Value.Trim() })
    $missing = @($RequiredSections | Where-Object { $headings -notcontains $_ })
    $detail = 'all present'
    if ($missing.Count) { $detail = 'missing: ' + ($missing -join ', ') }
    $checks.Add([pscustomobject]@{ Name = 'Required sections present'; Passed = ($missing.Count -eq 0); Detail = $detail })

    # A bracketed token with a letter, a digit and a hyphen counts as a case ID
    # citation, e.g. [SYN-2026-004]. Footnotes like [1] and links are ignored.
    $cited = @([regex]::Matches($Response, '\[([A-Za-z][A-Za-z0-9_./]*-[A-Za-z0-9_./-]*\d[A-Za-z0-9_./-]*)\](?!\()') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    $unknown = @($cited | Where-Object { $KnownCaseIds -notcontains $_ })
    $detail = "$($cited.Count) case IDs cited, all found in the input"
    if ($unknown.Count) { $detail = 'not in the input: ' + ($unknown -join ', ') }
    $checks.Add([pscustomobject]@{ Name = 'Cited case IDs exist'; Passed = ($unknown.Count -eq 0); Detail = $detail })
    $checks.ToArray()
}

# ---------------------------------------------------------------- docx (OpenXML, no Word needed)

function ConvertTo-XmlText {
    param([AllowEmptyString()][string]$Text)
    [Security.SecurityElement]::Escape($Text)
}

function New-WRun {
    param([string]$Text, [switch]$Bold, [string]$Color, [int]$HalfPoints)
    $props = ''
    if ($Bold) { $props += '<w:b/>' }
    if ($Color) { $props += "<w:color w:val=`"$Color`"/>" }
    if ($HalfPoints) { $props += "<w:sz w:val=`"$HalfPoints`"/>" }
    if ($props) { $props = "<w:rPr>$props</w:rPr>" }
    "<w:r>$props<w:t xml:space=`"preserve`">$(ConvertTo-XmlText $Text)</w:t></w:r>"
}

function New-WParagraph {
    <# A paragraph; **bold** spans in the text become bold runs. #>
    param([AllowEmptyString()][string]$Text, [string]$Style, [switch]$Bold, [string]$Color, [int]$Indent)
    $ppr = ''
    if ($Style) { $ppr += "<w:pStyle w:val=`"$Style`"/>" }
    if ($Indent) { $ppr += "<w:ind w:left=`"$Indent`" w:hanging=`"240`"/>" }
    if ($ppr) { $ppr = "<w:pPr>$ppr</w:pPr>" }
    $runs = ''
    $parts = [regex]::Split($Text, '(\*\*[^*]+\*\*)')
    foreach ($part in $parts) {
        if (-not $part) { continue }
        if ($part -match '^\*\*(.+)\*\*$') { $runs += New-WRun $Matches[1] -Bold -Color $Color }
        else { $runs += New-WRun $part -Bold:$Bold -Color $Color }
    }
    "<w:p>$ppr$runs</w:p>"
}

function New-WTable {
    <# $ColumnColor maps a column index to a text colour for the body rows, e.g. @{ 2 = '9E9E9E' }. #>
    param([string[]]$Header, [object[]]$Rows, [hashtable]$ColumnColor = @{})
    $cell = {
        param($text, $bold, $color)
        "<w:tc><w:tcPr><w:tcW w:w=`"0`" w:type=`"auto`"/></w:tcPr><w:p><w:pPr><w:pStyle w:val=`"TableText`"/></w:pPr>$(New-WRun ([string]$text) -Bold:$bold -Color $color)</w:p></w:tc>"
    }
    $xml = '<w:tbl><w:tblPr><w:tblStyle w:val="DossierTable"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
    $xml += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ((@($Header | ForEach-Object { & $cell $_ $true '' })) -join '') + '</w:tr>'
    foreach ($row in $Rows) {
        $cells = @($row)
        $xml += '<w:tr>' + ((@(for ($i = 0; $i -lt $cells.Count; $i++) { & $cell $cells[$i] $false ([string]$ColumnColor[$i]) })) -join '') + '</w:tr>'
    }
    $xml + '</w:tbl>'
}

function ConvertFrom-MarkdownToWordXml {
    <# Small Markdown subset: headings, bullets, numbered lists, bold, paragraphs. #>
    param([string]$Markdown)
    $Markdown = [regex]::Replace($Markdown, '(?s)<!--.*?-->', '')
    $out = New-Object System.Text.StringBuilder
    $paragraph = New-Object System.Collections.Generic.List[string]
    $flush = {
        if ($paragraph.Count) { [void]$out.Append((New-WParagraph ($paragraph -join ' '))); $paragraph.Clear() }
    }
    foreach ($line in ($Markdown -split "`r?`n")) {
        $trim = $line.Trim()
        if (-not $trim) { & $flush; continue }
        if ($trim -match '^(#{1,6})\s+(.+)$') {
            & $flush
            $style = 'Heading3'
            if ($Matches[1].Length -ge 3) { $style = 'Heading4' }
            [void]$out.Append((New-WParagraph $Matches[2] -Style $style))
        } elseif ($trim -match '^[-*]\s+(.+)$') {
            & $flush
            [void]$out.Append((New-WParagraph ([string][char]0x2022 + ' ' + $Matches[1]) -Indent 360))
        } elseif ($trim -match '^(\d+)[.)]\s+(.+)$') {
            & $flush
            [void]$out.Append((New-WParagraph ($Matches[1] + '. ' + $Matches[2]) -Indent 360))
        } elseif ($trim -match '^>\s?(.*)$') {
            & $flush
            [void]$out.Append((New-WParagraph $Matches[1] -Style 'Note'))
        } else { $paragraph.Add($trim) }
    }
    & $flush
    $out.ToString()
}

function Get-ParagraphText {
    param([string]$ParagraphXml)
    [Net.WebUtility]::HtmlDecode(([regex]::Replace($ParagraphXml, '<[^>]+>', '')))
}

function Set-AnchorParagraph {
    <#
      Replaces the paragraph whose visible text equals $Anchor. Word may split
      the anchor across several runs, so the match uses the paragraph text.
    #>
    param([string]$Xml, [string]$Anchor, [string]$Replacement)
    $pattern = '<w:p(?:\s[^>]*)?>(?:(?!</w:p>).)*?</w:p>'
    $result = [regex]::Replace($Xml, $pattern, {
            param($m)
            if ((Get-ParagraphText $m.Value).Trim() -eq $Anchor) { $script:anchorHits++; return $Replacement }
            $m.Value
        }, 'Singleline')
    $result
}

function Update-ZipEntries {
    <# Copies a zip, passing each entry's text through $Transform (name, text) -> text. #>
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
    } finally { $in.Dispose() }
    Move-Item -LiteralPath $tmp -Destination $Destination -Force
}

function New-DossierReport {
    <# Fills the report template: banner in the header, generated body at the anchor. #>
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][object[]]$Cases,
        [Parameter(Mandatory)]$Statistics,
        [Parameter(Mandatory)][string]$Response,
        [Parameter(Mandatory)][object[]]$Checks
    )
    $m = $script:Markers
    $body = New-Object System.Text.StringBuilder
    # The banner goes into the page header. Templates without a header anchor
    # get it as the first body paragraph instead.
    if (-not (Get-DocxText $TemplatePath).Contains('{{STATUS_BANNER}}')) {
        [void]$body.Append((New-WParagraph $m.Banner -Bold -Color 'C00000'))
    }
    [void]$body.Append((New-WParagraph $Title -Style 'Title'))
    [void]$body.Append((New-WTable @('Field', 'Value') @(
                , @('Run ID', $Run.RunId)
                , @('Generated (UTC)', $Run.GeneratedAt)
                , @('Mode', $Run.ModeText)
                , @('Model', $Run.Model)
                , @('Cases', "$($Statistics.Total) ($($Statistics.From) to $($Statistics.To))")
                , @('Workbook SHA-256', $Run.WorkbookSha256)
                , @('Prompt SHA-256', $Run.PromptSha256)
                , @('Response SHA-256', $Run.ResponseSha256)
            )))

    [void]$body.Append((New-WParagraph '1 Statistics' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph 'Cases by category' -Style 'Heading2'))
    [void]$body.Append((New-WTable @('Category', 'Count', 'Share %') @($Statistics.ByCategory | ForEach-Object { , @($_.Category, $_.Count, $_.Share.ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)) })))
    [void]$body.Append((New-WParagraph 'Cases by month' -Style 'Heading2'))
    $peak = ($Statistics.ByMonth | Measure-Object Count -Maximum).Maximum
    [void]$body.Append((New-WTable @('Month', 'Count', 'Trend') @($Statistics.ByMonth | ForEach-Object {
                    $bar = ''
                    if ($peak -gt 0) { $bar = ([string][char]0x2588) * [int][math]::Round(10.0 * $_.Count / $peak) }
                    , @($_.Month, $_.Count, $bar)
                }) -ColumnColor @{ 2 = '9E9E9E' }))

    [void]$body.Append((New-WParagraph '2 Case list' -Style 'Heading1'))
    [void]$body.Append((New-WTable @('Case ID', 'Received', 'Category', 'Severity', 'Status') @($Cases | ForEach-Object { , @($_.CaseId, $_.ReceivedDate.ToString('yyyy-MM-dd'), $_.Category, $_.Severity, $_.Status) })))

    [void]$body.Append((New-WParagraph '3 AI-drafted assessment' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph $Run.ResponseNote -Style 'Note'))
    [void]$body.Append((ConvertFrom-MarkdownToWordXml $Response))

    [void]$body.Append((New-WParagraph '4 Automated checks' -Style 'Heading1'))
    [void]$body.Append((New-WTable @('Check', 'Result', 'Detail') @($Checks | ForEach-Object {
                    $result = 'FAIL'
                    if ($_.Passed) { $result = 'pass' }
                    , @($_.Name, $result, $_.Detail)
                })))

    [void]$body.Append((New-WParagraph '5 Review and sign-off' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph 'This report is a draft until a reviewer has checked every statement in section 3 against the case records and signs it with Approve-ComplaintDossier.ps1. The signature and the file hashes are recorded in audit.json.'))
    [void]$body.Append((New-WParagraph 'Approval is refused while an automated check in section 4 has failed. A reviewer can accept failed checks only by giving a written justification, which is printed here and stored in audit.json as a documented exception.'))
    foreach ($key in 'Status', 'Reviewer', 'SignedAt', 'Comment', 'Exception') { [void]$body.Append((New-WParagraph $m[$key])) }

    $bodyXml = $body.ToString()
    $bannerXml = New-WParagraph $m.Banner -Bold -Color 'C00000'
    $script:anchorHits = 0
    $script:bannerHits = 0
    Update-ZipEntries -Source $TemplatePath -Destination $OutputPath -Transform {
        param($name, $text)
        if ($name -eq 'word/document.xml') { return Set-AnchorParagraph $text '{{REPORT_BODY}}' $bodyXml }
        if ($name -match '^word/header\d*\.xml$') {
            $before = $script:anchorHits
            $text = Set-AnchorParagraph $text '{{STATUS_BANNER}}' $bannerXml
            $script:bannerHits += $script:anchorHits - $before
            $script:anchorHits = $before
            return $text
        }
        $text
    }
    if ($script:anchorHits -ne 1) {
        Remove-Item -LiteralPath $OutputPath -Force
        throw "Report template must contain exactly one paragraph with the text {{REPORT_BODY}} (found $script:anchorHits)."
    }
}

function Get-DocxText {
    <# Visible text of document.xml and headers, one paragraph per line. For tests and checks. #>
    param([Parameter(Mandatory)][string]$Path)
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $parts = @($zip.Entries | Where-Object { $_.FullName -match '^word/(document|header\d*)\.xml$' } | Sort-Object FullName)
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

# ---------------------------------------------------------------- live Copilot handoff

function Invoke-CopilotHandoff {
    <#
      Runs src/Copilot/Invoke-CopilotRoundtrip.ps1 in Windows PowerShell 5.1
      (STA, .NET Framework UI Automation). The child opens a new chat, selects
      the configured model (or aborts), uploads the files, sends the prompt,
      waits for the answer and copies it into $ResponseFile.
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
    if ([string](Get-ConfigValue $cp 'PreferredModelGroup')) { $arguments += @('-PreferredModelGroup', [string]$cp.PreferredModelGroup) }
    if ([bool](Get-ConfigValue $cp 'TemporaryChat')) { $arguments += '-TemporaryChat' }
    foreach ($key in 'StartDelaySeconds', 'NewChatTimeoutSeconds', 'UploadTimeoutSeconds', 'ResponseTimeoutSeconds',
        'StableReadyChecks', 'PollMilliseconds', 'ModelMenuTimeoutMilliseconds') {
        $value = Get-ConfigValue $cp $key
        if ($null -ne $value) { $arguments += @("-$key", [string][int]$value) }
    }
    & $Executable @arguments
    $exit = $LASTEXITCODE
    $status = $null
    if (Test-Path -LiteralPath $statusFile) { $status = Read-JsonFile $statusFile }
    if ($exit -ne 0) {
        $reason = 'see the output above'
        if ($null -ne $status -and $status.error) { $reason = $status.error }
        throw "Copilot handoff failed (exit code $exit): $reason"
    }
    if (-not (Test-Path -LiteralPath $ResponseFile) -or [string]::IsNullOrWhiteSpace((Read-Utf8File $ResponseFile))) {
        throw 'Copilot handoff ended without a response.'
    }
    $status
}

# ---------------------------------------------------------------- audit and approval

function Save-AuditLog {
    param([Parameter(Mandatory)]$Audit, [Parameter(Mandatory)][string]$Path)
    Write-Utf8File $Path (($Audit | ConvertTo-Json -Depth 10) + "`n")
}

function Approve-DossierRun {
    <#
      Human sign-off. Verifies that prompt, response and draft report still match
      the hashes in audit.json, writes a signed copy of the report without the
      DRAFT markers, and records reviewer, time and new hash. The draft stays.

      Approval is refused while any automated response check in audit.json has
      failed. -AcceptFailedChecks with a non-empty -Justification overrides that;
      the override is printed in the signed report and stored in audit.json as a
      documented exception.
    #>
    param(
        [Parameter(Mandatory)][string]$RunPath,
        [Parameter(Mandatory)][string]$Reviewer,
        [string]$Comment = '',
        [switch]$AcceptFailedChecks,
        [string]$Justification = ''
    )
    if ([string]::IsNullOrWhiteSpace($Reviewer)) { throw 'Reviewer name is required.' }
    if ($AcceptFailedChecks -and [string]::IsNullOrWhiteSpace($Justification)) {
        throw '-AcceptFailedChecks needs a -Justification that says why the failed checks are acceptable. Approval refused.'
    }
    if (-not $AcceptFailedChecks -and -not [string]::IsNullOrWhiteSpace($Justification)) {
        throw '-Justification is only used together with -AcceptFailedChecks. Approval refused.'
    }
    $auditPath = Join-Path $RunPath 'audit.json'
    if (-not (Test-Path -LiteralPath $auditPath)) { throw "audit.json not found in $RunPath" }
    $audit = Read-JsonFile $auditPath
    if ($audit.status -ne 'DRAFT') { throw "Run $($audit.runId) has status $($audit.status); only DRAFT runs can be approved." }
    if ($audit.PSObject.Properties.Name -notcontains 'checks' -or @($audit.checks).Count -eq 0) {
        throw "audit.json of run $($audit.runId) records no automated checks. Approval refused."
    }
    $failed = @($audit.checks | Where-Object { $_.passed -ne $true })
    $failedText = @($failed | ForEach-Object {
            if ($_.detail) { "$($_.name) ($($_.detail))" } else { [string]$_.name }
        }) -join '; '
    if ($failed.Count -and -not $AcceptFailedChecks) {
        throw ("Approval refused: $($failed.Count) automated check(s) failed: $failedText. The run stays DRAFT. " +
            "Start a new run, or accept the failures as a documented exception with -AcceptFailedChecks -Justification '<reason>'.")
    }
    if ($AcceptFailedChecks -and -not $failed.Count) {
        throw 'All automated checks passed; -AcceptFailedChecks is not needed. Approval refused.'
    }

    foreach ($name in 'prompt', 'response', 'statistics', 'report') {
        $artifact = $audit.artifacts.$name
        $file = Join-Path $RunPath $artifact.file
        if (-not (Test-Path -LiteralPath $file)) { throw "Artifact missing: $($artifact.file). Approval refused." }
        $actual = Get-Sha256 $file
        if ($actual -ne $artifact.sha256) { throw "Hash mismatch for $($artifact.file): audit has $($artifact.sha256), file has $actual. The file changed after the run. Approval refused." }
    }

    $signedAt = Get-UtcStamp
    $m = $script:Markers
    $replace = [ordered]@{
        $m.Banner   = "REVIEWED AND APPROVED by $Reviewer on $signedAt"
        $m.Status   = 'Status: APPROVED'
        $m.Reviewer = "Reviewer: $Reviewer"
        $m.SignedAt = "Signed at (UTC): $signedAt"
    }
    $commentText = 'Review comment: none given'
    if ($Comment) { $commentText = "Review comment: $Comment" }
    $replace[$m.Comment] = $commentText
    $exception = $null
    if ($AcceptFailedChecks) {
        $exception = [pscustomobject]@{
            acceptedFailedChecks = @($failed | ForEach-Object { [pscustomobject]@{ name = $_.name; detail = $_.detail } })
            justification        = $Justification.Trim()
        }
        $replace[$m.Banner] = "REVIEWED AND APPROVED WITH A DOCUMENTED EXCEPTION by $Reviewer on $signedAt"
        $replace[$m.Status] = 'Status: APPROVED with a documented exception (failed automated checks accepted)'
        $replace[$m.Exception] = "Accepted check failures: DOCUMENTED EXCEPTION. Failed checks: $failedText. Justification: $($exception.justification)"
    } else {
        $replace[$m.Exception] = 'Accepted check failures: none (all automated checks passed)'
    }

    $draft = Join-Path $RunPath $audit.artifacts.report.file
    $approvedName = $audit.artifacts.report.file -replace '-DRAFT\.docx$', '.docx'
    if ($approvedName -eq $audit.artifacts.report.file) { $approvedName = [IO.Path]::GetFileNameWithoutExtension($approvedName) + '-approved.docx' }
    $approved = Join-Path $RunPath $approvedName
    $script:markerHits = @{}
    Update-ZipEntries -Source $draft -Destination $approved -Transform {
        param($name, $text)
        if ($name -notmatch '^word/(document|header\d*)\.xml$') { return $text }
        foreach ($key in $replace.Keys) {
            $old = ConvertTo-XmlText $key
            $hits = ([regex]::Matches($text, [regex]::Escape($old))).Count
            if ($hits) {
                $script:markerHits[$key] = [int]$script:markerHits[$key] + $hits
                $text = $text.Replace($old, (ConvertTo-XmlText $replace[$key]))
            }
        }
        $text
    }
    foreach ($key in $replace.Keys) {
        if (-not $script:markerHits[$key]) {
            Remove-Item -LiteralPath $approved -Force
            throw "Draft report does not contain the marker '$key'. Approval refused."
        }
    }
    $approvedText = Get-DocxText $approved
    foreach ($key in $replace.Keys) {
        if ($approvedText.Contains($key)) {
            Remove-Item -LiteralPath $approved -Force
            throw "Approved report still contains '$key'. Approval refused."
        }
    }

    $audit.status = 'APPROVED'
    $audit.review = [pscustomobject]@{
        reviewer        = $Reviewer
        signedAtUtc     = $signedAt
        comment         = $Comment
        osAccount       = [Environment]::UserName
        draftSha256     = $audit.artifacts.report.sha256
        exception       = $exception
    }
    $audit.artifacts | Add-Member -NotePropertyName approvedReport -NotePropertyValue ([pscustomobject]@{ file = $approvedName; sha256 = (Get-Sha256 $approved) }) -Force
    if ($exception) {
        $audit.events += [pscustomobject]@{ atUtc = $signedAt; event = 'approved-with-exception'; detail = "by $Reviewer; accepted failed checks: $failedText" }
    } else {
        $audit.events += [pscustomobject]@{ atUtc = $signedAt; event = 'approved'; detail = "by $Reviewer" }
    }
    Save-AuditLog $audit $auditPath
    [pscustomobject]@{ RunId = $audit.runId; Report = $approved; Sha256 = $audit.artifacts.approvedReport.sha256; SignedAt = $signedAt; Exception = $exception }
}

Export-ModuleMember -Function *-*
