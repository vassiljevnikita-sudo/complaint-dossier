# Complaint Dossier core functions.
# Drafts one complaint report per case: complaint list (Excel) -> case selection ->
# per-case prompt -> Microsoft 365 Copilot (or a prewritten demo answer) ->
# DRAFT Word report -> human review and approval, with an audit log per report.
# Runs on PowerShell 7 (Windows, Linux, macOS) and on Windows PowerShell 5.1.
# Keep this file ASCII-only: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.

Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:ToolVersion = '0.3.0'
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:FixedZipTime = [DateTimeOffset]::new(2026, 1, 1, 0, 0, 0, [TimeSpan]::Zero)

# Text markers that Approve-ComplaintDossier replaces. Each must appear in the draft.
$script:Markers = [ordered]@{
    Banner     = 'DRAFT - NOT REVIEWED - NOT FOR USE'
    Status     = 'Status: DRAFT (awaiting human review)'
    Reviewer   = 'Approved by: (not approved)'
    ApprovedAt = 'Approved at (UTC): (not approved)'
    Comment    = 'Review comment: (none)'
    Exception  = 'Accepted check failures: (not approved)'
}

# Headings of the data part of every prompt. A response that contains one of
# them is a copy of the prompt, not an answer.
$script:PromptHeadings = @(
    '### Complaint record (from the complaint list)',
    '### Similar complaints (calculated by the tool)',
    '### Case documents',
    '### Attached files'
)

# Neutral case schema. Workbook columns are mapped to these names in Workbook.Columns;
# only CaseId, ReceivedDate, Category and Description are required.
$script:CaseFields = @('CaseId', 'ReceivedDate', 'Category', 'ProductGroup', 'Batch', 'Severity', 'Status', 'SampleReceived', 'Description')
$script:ReportTypes = @('Short', 'Long')

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
        @('CaseFolders', 'Path'), @('Reports', 'Short'), @('Reports', 'Long'),
        @('Report', 'TemplatePath'), @('Report', 'OutputPath')
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
    $prompts = [ordered]@{}
    foreach ($type in $script:ReportTypes) {
        $settings = $config.Reports.$type
        foreach ($key in 'Marker', 'PromptPath', 'Title') {
            if ([string]::IsNullOrWhiteSpace([string](Get-ConfigValue $settings $key))) { throw "Configuration value missing: Reports.$type.$key" }
        }
        if (@(@(Get-ConfigValue $settings 'RequiredSections') | Where-Object { $_ }).Count -eq 0) { throw "Configuration value missing: Reports.$type.RequiredSections" }
        $prompts[$type] = Resolve-DossierPath $Root $settings.PromptPath
    }
    if ($config.Reports.Short.Marker -eq $config.Reports.Long.Marker) { throw 'Reports.Short.Marker and Reports.Long.Marker must differ.' }

    $paths = [ordered]@{
        Config         = [IO.Path]::GetFullPath($Path)
        Workbook       = Resolve-DossierPath $Root $config.Workbook.Path
        CaseFolders    = Resolve-DossierPath $Root $config.CaseFolders.Path
        Prompts        = [pscustomobject]$prompts
        ReportTemplate = Resolve-DossierPath $Root $config.Report.TemplatePath
        Output         = Resolve-DossierPath $Root $config.Report.OutputPath
        DemoResponses  = Resolve-DossierPath $Root ([string](Get-ConfigValue $config.Demo 'ResponsePath'))
        Upload         = Resolve-DossierPath $Root ([string](Get-ConfigValue (Get-ConfigValue $config.Copilot 'Upload') 'Path'))
    }
    foreach ($key in 'Workbook', 'ReportTemplate') {
        if (-not (Test-Path -LiteralPath $paths[$key] -PathType Leaf)) { throw "Configured file not found ($key): $($paths[$key])" }
    }
    foreach ($type in $script:ReportTypes) {
        if (-not (Test-Path -LiteralPath $prompts[$type] -PathType Leaf)) { throw "Configured file not found (Reports.$type.PromptPath): $($prompts[$type])" }
    }
    if (-not (Test-Path -LiteralPath $paths.CaseFolders -PathType Container)) { throw "Configured case folder root not found (CaseFolders.Path): $($paths.CaseFolders)" }
    if ($Mode -eq 'Demo' -and -not (Test-Path -LiteralPath $paths.DemoResponses -PathType Container)) {
        throw "Configured folder not found (Demo.ResponsePath): $($paths.DemoResponses)"
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
    <# Reads the complaint list and maps workbook columns to the neutral case schema. All rows, sorted by date. #>
    param([Parameter(Mandatory)]$Config)
    $wb = $Config.Settings.Workbook
    $headerRow = 1
    if ($null -ne (Get-ConfigValue $wb 'HeaderRow')) { $headerRow = [int]$wb.HeaderRow }
    $rows = @(Read-XlsxSheet -Path $Config.Paths.Workbook -Sheet $wb.Sheet -HeaderRow $headerRow)
    if ($rows.Count -eq 0) { throw "Worksheet '$($wb.Sheet)' contains no cases." }

    $map = [ordered]@{}
    foreach ($field in $script:CaseFields) {
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
        if ($id -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { throw "Row $($row._Row): case ID '$id' may only contain letters, digits, '.', '_' and '-', because it names the case folder." }
        $case = [ordered]@{}
        foreach ($field in $script:CaseFields) {
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
    @($cases | Sort-Object ReceivedDate, CaseId)
}

# ---------------------------------------------------------------- case selection

function Get-CaseFolder {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    Join-Path $Config.Paths.CaseFolders $CaseId
}

function Get-ReportTypeMarker {
    <#
      The report type of a case comes from an empty marker file in its case
      folder (by default short.txt or long.txt). Returns 'Short', 'Long' or $null.
      Both markers at once is an error.
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $folder = Get-CaseFolder $Config $CaseId
    $found = @(foreach ($type in $script:ReportTypes) {
            if (Test-Path -LiteralPath (Join-Path $folder $Config.Settings.Reports.$type.Marker) -PathType Leaf) { $type }
        })
    if ($found.Count -gt 1) {
        throw "Case folder of $CaseId holds both $($Config.Settings.Reports.Short.Marker) and $($Config.Settings.Reports.Long.Marker). Keep exactly one."
    }
    if ($found.Count -eq 1) { return $found[0] }
    $null
}

function Find-ApprovedReport {
    <# Path of an approved report for this case under Report.OutputPath, or $null. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $out = $Config.Paths.Output
    if (-not (Test-Path -LiteralPath $out -PathType Container)) { return $null }
    foreach ($run in @(Get-ChildItem -LiteralPath $out -Directory | Sort-Object Name)) {
        $audit = Join-Path (Join-Path $run.FullName $CaseId) 'audit.json'
        if (-not (Test-Path -LiteralPath $audit -PathType Leaf)) { continue }
        $status = $null
        try { $status = (Read-JsonFile $audit).status } catch { $status = $null }
        if ($status -eq 'APPROVED') { return (Split-Path -Parent $audit) }
    }
    $null
}

function Select-DossierCases {
    <#
      Decides which cases get a report in this run.
      Without -CaseId: every case whose status is listed in Selection.OpenStatuses.
      With -CaseId: exactly these cases, whatever their status.
      A selected case is skipped when its case folder has no report type marker
      or when an approved report for it already exists. Returns one entry per
      selected case with Type and SkipReason (empty when the case is drafted).
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][object[]]$Cases, [string[]]$CaseId)
    if ($CaseId) {
        $unknown = @($CaseId | Where-Object { @($Cases | ForEach-Object { $_.CaseId }) -notcontains $_ })
        if ($unknown.Count) { throw "Case ID not found in the complaint list: $($unknown -join ', ')" }
        $chosen = @($Cases | Where-Object { $CaseId -contains $_.CaseId })
    } else {
        $open = @(@(Get-ConfigValue (Get-ConfigValue $Config.Settings 'Selection') 'OpenStatuses') | Where-Object { $_ })
        if ($open.Count -eq 0) { $open = @('Open') }
        $chosen = @($Cases | Where-Object { $open -contains $_.Status })
    }
    foreach ($case in $chosen) {
        $type = $null
        $reason = ''
        try { $type = Get-ReportTypeMarker -Config $Config -CaseId $case.CaseId } catch { $reason = $_.Exception.Message }
        if (-not $reason -and -not $type) {
            $reason = "no report type marker ($($Config.Settings.Reports.Short.Marker) or $($Config.Settings.Reports.Long.Marker)) in the case folder"
        }
        if (-not $reason) {
            $approved = Find-ApprovedReport -Config $Config -CaseId $case.CaseId
            if ($approved) { $reason = "an approved report already exists: $(Get-RelativeDisplayPath $Config.Root $approved -FullPathOutside)" }
        }
        [pscustomobject]@{ Case = $case; Type = $type; SkipReason = $reason }
    }
}

# ---------------------------------------------------------------- similar complaints

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
    @($script:ReportTypes | ForEach-Object { ([string]$Config.Settings.Reports.$_.Marker).ToLowerInvariant() })
}

function Get-CaseDocuments {
    <# Text files from the case folder whose content goes into the prompt. Marker files are left out. #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $folder = Get-CaseFolder $Config $CaseId
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return @() }
    $settings = $Config.Settings.CaseFolders
    $extensions = @(@(Get-ConfigValue $settings 'TextExtensions') | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
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
      Files that the live run attaches to the Copilot chat for one case:
      <Copilot.Upload.Path>/<case id>/, or the case folder when Upload.Path is empty.
      Only Copilot.Upload.Extensions; marker files are left out. Fails closed on
      empty files, oversize files or too many files.
    #>
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$CaseId)
    $dir = Get-CaseFolder $Config $CaseId
    if ($Config.Paths.Upload) { $dir = Join-Path $Config.Paths.Upload $CaseId }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { return @() }
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
      human review; the results are printed in the report and stored in audit.json.
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
    $checks.Add([pscustomobject]@{ Name = 'Response is not empty'; Passed = -not [string]::IsNullOrWhiteSpace($Response); Detail = "$($Response.Length) characters" })
    $echo = (Get-TextSha256 $Response.Trim()) -eq (Get-TextSha256 ([string]$Prompt).Trim())
    $leaked = @($script:PromptHeadings | Where-Object { $Response.Contains($_) })
    $detail = ''
    if ($echo) { $detail = 'response equals the prompt' } elseif ($leaked.Count) { $detail = 'contains prompt heading: ' + ($leaked -join ', ') }
    $checks.Add([pscustomobject]@{ Name = 'Response is not a copy of the prompt'; Passed = (-not $echo -and $leaked.Count -eq 0); Detail = $detail })

    $headings = @([regex]::Matches($Response, '(?m)^#{1,6}\s+(.+?)\s*$') | ForEach-Object { $_.Groups[1].Value.Trim() })
    $missing = @($RequiredSections | Where-Object { $headings -notcontains $_ })
    $detail = 'all present'
    if ($missing.Count) { $detail = 'missing: ' + ($missing -join ', ') }
    $checks.Add([pscustomobject]@{ Name = 'Required sections present'; Passed = ($missing.Count -eq 0); Detail = $detail })

    # A bracketed token with a letter, a digit and a hyphen counts as a case ID
    # citation, e.g. [SYN-2026-004]. Footnotes like [1] and links are ignored.
    $cited = @([regex]::Matches($Response, '\[([A-Za-z][A-Za-z0-9_./]*-[A-Za-z0-9_./-]*\d[A-Za-z0-9_./-]*)\](?!\()') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    # The report must name its own case, cited in brackets or written plainly
    # (for example in the basic data list "Case ID: SYN-2026-007").
    $own = $cited -contains $CaseId
    $detail = "[$CaseId] cited"
    if (-not $own -and $CaseId -and [regex]::IsMatch($Response, '(?<![A-Za-z0-9_-])' + [regex]::Escape($CaseId) + '(?![A-Za-z0-9_-])')) { $own = $true; $detail = "$CaseId named" }
    if (-not $own) { $detail = "$CaseId not named" }
    $checks.Add([pscustomobject]@{ Name = 'Report names its own case ID'; Passed = $own; Detail = $detail })
    $unknown = @($cited | Where-Object { $KnownCaseIds -notcontains $_ })
    $detail = "$($cited.Count) case IDs cited, all in the complaint list"
    if ($unknown.Count) { $detail = 'not in the complaint list: ' + ($unknown -join ', ') }
    $checks.Add([pscustomobject]@{ Name = 'Cited case IDs exist'; Passed = ($unknown.Count -eq 0); Detail = $detail })

    $absent = @($RequiredStatements | Where-Object { $Response.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -lt 0 })
    $detail = 'none required for this case'
    if (@($RequiredStatements).Count) { $detail = 'present: ' + ($RequiredStatements -join ' ') }
    if ($absent.Count) { $detail = 'missing: ' + ($absent -join ' ') }
    $checks.Add([pscustomobject]@{ Name = 'Required statements present'; Passed = ($absent.Count -eq 0); Detail = $detail })

    if (@($ClassificationStatements).Count) {
        $end = $Response.TrimEnd()
        $match = @($ClassificationStatements | Where-Object { $end.EndsWith($_, [StringComparison]::OrdinalIgnoreCase) })
        $detail = 'ends with: ' + ($match -join '')
        if (-not $match.Count) { $detail = 'last sentence is none of: ' + ($ClassificationStatements -join ' / ') }
        $checks.Add([pscustomobject]@{ Name = 'Conclusion ends with a classification'; Passed = ($match.Count -gt 0); Detail = $detail })
    }
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
        } elseif ($trim -match ('^[-*' + [char]0x2022 + ']\s+(.+)$')) {
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

function New-CaseReport {
    <# Fills the report template for one case: banner in the header, generated body at the anchor. #>
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)]$Case,
        [Parameter(Mandatory)]$Similar,
        [object[]]$Documents = @(),
        [string[]]$Uploads = @(),
        [Parameter(Mandatory)][string]$Response,
        [Parameter(Mandatory)][object[]]$Checks,
        [bool]$Mask = $true
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
                , @('Case ID', $Case.CaseId)
                , @('Report type', $Run.ReportType)
                , @('Run ID', $Run.RunId)
                , @('Generated (UTC)', $Run.GeneratedAt)
                , @('Mode', $Run.ModeText)
                , @('Model', $Run.Model)
                , @('Complaint list SHA-256', $Run.WorkbookSha256)
                , @('Prompt SHA-256', $Run.PromptSha256)
                , @('Response SHA-256', $Run.ResponseSha256)
            )))

    [void]$body.Append((New-WParagraph '1 Complaint record' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph 'From the complaint list. Contact data in the description is masked.' -Style 'Note'))
    [void]$body.Append((New-WTable @('Field', 'Value') @(Get-CaseRecordRows $Case $Mask | Select-Object -Skip 1)))

    [void]$body.Append((New-WParagraph '2 Similar complaints' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph 'Computed by the tool from the complaint list, not by the model.' -Style 'Note'))
    foreach ($line in Get-SimilarComplaintSentences $Case $Similar) { [void]$body.Append((New-WParagraph ([string][char]0x2022 + ' ' + $line) -Indent 360)) }
    $listed = @(@($Similar.SameBatch) + @($Similar.SameProductGroupCategory) | Sort-Object CaseId -Unique)
    if ($listed.Count) {
        [void]$body.Append((New-WTable @('Case ID', 'Received', 'Category', 'Product group', 'Batch', 'Status') @($listed | ForEach-Object {
                        , @($_.CaseId, $_.ReceivedDate, $_.Category, $_.ProductGroup, $_.Batch, $_.Status)
                    })))
    }

    [void]$body.Append((New-WParagraph '3 Case documents' -Style 'Heading1'))
    $rows = @(@($Documents | ForEach-Object { , @($_.Name, 'text inserted into the prompt') }) + @($Uploads | ForEach-Object { , @($_, 'uploaded to the Copilot chat') }))
    if ($rows.Count) { [void]$body.Append((New-WTable @('File', 'Use') $rows)) }
    else { [void]$body.Append((New-WParagraph 'The case folder holds no documents for the prompt and no files to upload.')) }

    [void]$body.Append((New-WParagraph '4 Report text (AI draft)' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph $Run.ResponseNote -Style 'Note'))
    [void]$body.Append((ConvertFrom-MarkdownToWordXml $Response))

    [void]$body.Append((New-WParagraph '5 Automated checks' -Style 'Heading1'))
    [void]$body.Append((New-WTable @('Check', 'Result', 'Detail') @($Checks | ForEach-Object {
                    $result = 'FAIL'
                    if ($_.Passed) { $result = 'pass' }
                    , @($_.Name, $result, $_.Detail)
                })))

    [void]$body.Append((New-WParagraph '6 Review and approval' -Style 'Heading1'))
    [void]$body.Append((New-WParagraph "This report is a draft until a reviewer has checked every statement in section 4 against the case file of $($Case.CaseId) and approves it with Approve-ComplaintDossier.ps1. The approval (typed reviewer name and time) and the file hashes are recorded in the audit.json of this report. A typed name is not an electronic signature."))
    [void]$body.Append((New-WParagraph 'Approval is refused while an automated check in section 5 has failed. A reviewer can accept failed checks only by giving a written justification, which is printed here and stored in audit.json as a documented exception.'))
    foreach ($key in 'Status', 'Reviewer', 'ApprovedAt', 'Comment', 'Exception') { [void]$body.Append((New-WParagraph $m[$key])) }

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

# ---------------------------------------------------------------- one report per case

function Get-ReportFileName {
    param([Parameter(Mandatory)][string]$CaseId, [Parameter(Mandatory)][string]$Type)
    "$CaseId-$($Type.ToLowerInvariant())-report-DRAFT.docx"
}

function New-DossierCaseReport {
    <#
      Drafts the report for one selected case into <RunPath>/<case id>/:
      prompt.txt, case-data.json, response.md, the DRAFT .docx and audit.json.
      Demo mode copies <Demo.ResponsePath>/<case id>.md instead of asking Copilot.
    #>
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)]$Selected,
        [Parameter(Mandatory)][object[]]$AllCases,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$RunPath
    )
    $case = $Selected.Case
    $type = $Selected.Type
    $settings = $Config.Settings.Reports.$type
    $root = $Config.Root
    $casePath = Join-Path $RunPath $case.CaseId
    New-Item -ItemType Directory -Force -Path $casePath | Out-Null

    $events = New-Object System.Collections.Generic.List[object]
    $addEvent = { param($name, $detail) $events.Add([pscustomobject]@{ atUtc = (Get-UtcStamp); event = $name; detail = $detail }) }
    $inputs = New-Object System.Collections.Generic.List[object]
    $addInput = { param($role, $path) $inputs.Add([pscustomobject]@{ role = $role; path = (Get-RelativeDisplayPath $root $path); sha256 = (Get-Sha256 $path) }) }
    & $addEvent 'case-started' "$($case.CaseId), $($type.ToLowerInvariant()) report, mode=$($Config.Mode.ToLowerInvariant())"
    & $addInput 'config' $Config.Paths.Config
    & $addInput 'complaint-list' $Config.Paths.Workbook
    & $addInput 'prompt-template' $Config.Paths.Prompts.$type
    & $addInput 'report-template' $Config.Paths.ReportTemplate

    # 1. Case data: row, similar complaints, documents, uploads
    $lookback = 12
    $recentDays = 30
    $similarSettings = Get-ConfigValue $Config.Settings 'SimilarComplaints'
    if ($null -ne (Get-ConfigValue $similarSettings 'LookbackMonths')) { $lookback = [int]$similarSettings.LookbackMonths }
    if ($null -ne (Get-ConfigValue $similarSettings 'RecentDays')) { $recentDays = [int]$similarSettings.RecentDays }
    $levelArgs = @{}
    if (@(@(Get-ConfigValue $similarSettings 'FrequencyLevels') | Where-Object { $_ }).Count) { $levelArgs.FrequencyLevels = @($similarSettings.FrequencyLevels) }
    $similar = Get-SimilarComplaints -Case $case -Cases $AllCases -LookbackMonths $lookback -RecentDays $recentDays @levelArgs
    $documents = @(Get-CaseDocuments -Config $Config -CaseId $case.CaseId)
    foreach ($doc in $documents) { & $addInput 'case-document' $doc.Path }
    $uploads = @()
    if ($Config.Mode -eq 'Live') {
        $uploads = @(Get-UploadFiles -Config $Config -CaseId $case.CaseId)
        foreach ($file in $uploads) { & $addInput 'upload' $file.FullName }
    }
    $statements = @(Get-RequiredStatements -Config $Config -Case $case -Type $type -Similar $similar)
    $facts = @(@(Get-ConfigValue $settings 'ProcessFacts') | Where-Object { $_ })
    $caseDataFile = Join-Path $casePath 'case-data.json'
    $caseData = [pscustomobject][ordered]@{
        caseId             = $case.CaseId
        reportType         = $type.ToLowerInvariant()
        record             = [pscustomobject][ordered]@{
            CaseId = $case.CaseId; ReceivedDate = $case.ReceivedDate.ToString('yyyy-MM-dd'); Category = $case.Category
            ProductGroup = $case.ProductGroup; Batch = $case.Batch; Severity = $case.Severity; Status = $case.Status
            SampleReceived = $case.SampleReceived; Description = $case.Description
        }
        similarComplaints  = $similar
        requiredStatements = $statements
        processFacts       = $facts
        documents          = @($documents | ForEach-Object { $_.Name })
        uploads            = @($uploads | ForEach-Object { $_.Name })
    }
    # JSON with LF line ends, so the hash is the same on every platform.
    Write-Utf8File $caseDataFile (($caseData | ConvertTo-Json -Depth 6).Replace("`r`n", "`n") + "`n")

    # 2. Prompt
    $mask = $true
    if ($null -ne (Get-ConfigValue (Get-ConfigValue $Config.Settings 'Prompt') 'MaskContactData')) { $mask = [bool]$Config.Settings.Prompt.MaskContactData }
    $prompt = New-CasePrompt -Template (Read-Utf8File $Config.Paths.Prompts.$type) -Case $case -Similar $similar -Documents $documents `
        -Attachments @($uploads | ForEach-Object { $_.Name }) -RequiredStatements $statements -ProcessFacts $facts -Mask $mask
    $promptFile = Join-Path $casePath 'prompt.txt'
    Write-Utf8File $promptFile $prompt
    & $addEvent 'prompt-built' "$($prompt.Length) characters, $($documents.Count) document(s), $($uploads.Count) upload(s)"

    # 3. Response
    $responseFile = Join-Path $casePath 'response.md'
    $copilotStatus = $null
    if ($Config.Mode -eq 'Demo') {
        $demoFile = Join-Path $Config.Paths.DemoResponses "$($case.CaseId).md"
        if (-not (Test-Path -LiteralPath $demoFile -PathType Leaf)) { throw "No prewritten demo answer for $($case.CaseId): $(Get-RelativeDisplayPath $root $demoFile -FullPathOutside)" }
        & $addInput 'demo-response' $demoFile
        Write-Utf8File $responseFile (Read-Utf8File $demoFile)
        $modelText = 'none (demo mode)'
        $modeText = 'Demo: prewritten sample answer, Copilot not used'
        $responseNote = "Demo mode: the text below is a prewritten sample answer for the synthetic case $($case.CaseId). It was not produced by Copilot."
        & $addEvent 'response-received' 'demo answer copied'
    } else {
        $cp = $Config.Settings.Copilot
        $modelText = [string]$cp.PreferredModel
        if ([string](Get-ConfigValue $cp 'PreferredModelGroup')) { $modelText = "$($cp.PreferredModelGroup) / $modelText" }
        $modeText = 'Live: Microsoft 365 Copilot desktop app'
        $responseNote = "Drafted by Microsoft 365 Copilot (configured model: $modelText). Check every statement against the case file."
        Write-Host "[copilot] $($case.CaseId): handing the prompt to '$($cp.WindowTitle)', model '$modelText', $($uploads.Count) upload(s)"
        $copilotStatus = Invoke-CopilotHandoff -Config $Config -PromptFile $promptFile -ResponseFile $responseFile -UploadFiles $uploads -RunPath $casePath
        & $addEvent 'response-received' "copilot, model=$modelText, uploads=$($uploads.Count)"
    }
    # Plain-text answers (Copilot copy button) get their section headings back.
    $response = ConvertTo-SectionHeadings (Read-Utf8File $responseFile) @($settings.RequiredSections)

    # 4. Checks and report
    $checks = @(Test-ReportResponse -Response $response -Prompt $prompt -CaseId $case.CaseId -RequiredSections @($settings.RequiredSections) `
            -KnownCaseIds @($AllCases | ForEach-Object { $_.CaseId }) -RequiredStatements $statements `
            -ClassificationStatements @(@(Get-ConfigValue $settings 'ClassificationStatements') | Where-Object { $_ }))
    $reportName = Get-ReportFileName $case.CaseId $type
    $reportFile = Join-Path $casePath $reportName
    $run = [pscustomobject]@{
        RunId          = $RunId
        ReportType     = "$type report"
        GeneratedAt    = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss')
        ModeText       = $modeText
        Model          = $modelText
        ResponseNote   = $responseNote
        WorkbookSha256 = (Get-Sha256 $Config.Paths.Workbook)
        PromptSha256   = (Get-Sha256 $promptFile)
        ResponseSha256 = (Get-Sha256 $responseFile)
    }
    New-CaseReport -TemplatePath $Config.Paths.ReportTemplate -OutputPath $reportFile -Title "$($settings.Title) $($case.CaseId)" -Run $run `
        -Case $case -Similar $similar -Documents $documents -Uploads @($uploads | ForEach-Object { $_.Name }) -Response $response -Checks $checks -Mask $mask
    & $addEvent 'report-written' $reportName

    # 5. Audit log of this report (SHA-256 hashes)
    $artifact = { param($name) [pscustomobject]@{ file = $name; sha256 = (Get-Sha256 (Join-Path $casePath $name)) } }
    $audit = [pscustomobject][ordered]@{
        schema     = 'complaint-dossier/audit/v3'
        tool       = [pscustomobject]@{ name = 'Complaint Dossier'; version = (Get-DossierVersion) }
        runId      = $RunId
        caseId     = $case.CaseId
        reportType = $type.ToLowerInvariant()
        mode       = $Config.Mode.ToLowerInvariant()
        status     = 'DRAFT'
        operator   = [pscustomobject]@{ osAccount = [Environment]::UserName; host = 'PowerShell ' + $PSVersionTable.PSVersion.ToString() + ' (' + $PSVersionTable.PSEdition + ')' }
        model      = [pscustomobject]@{ configured = $modelText; copilot = $copilotStatus }
        inputs     = $inputs.ToArray()
        artifacts  = [pscustomobject][ordered]@{
            caseData = & $artifact 'case-data.json'
            prompt   = & $artifact 'prompt.txt'
            response = & $artifact 'response.md'
            report   = & $artifact $reportName
        }
        checks     = @($checks | ForEach-Object { [pscustomobject]@{ name = $_.Name; passed = $_.Passed; detail = $_.Detail } })
        review     = [pscustomobject]@{ reviewer = $null; approvedAtUtc = $null; comment = $null }
        events     = $events.ToArray()
    }
    $auditFile = Join-Path $casePath 'audit.json'
    Save-AuditLog $audit $auditFile
    [pscustomobject]@{ CaseId = $case.CaseId; Type = $type; Path = $casePath; Report = $reportFile; Audit = $auditFile; Checks = $checks }
}

# ---------------------------------------------------------------- audit and approval

function Save-AuditLog {
    param([Parameter(Mandatory)]$Audit, [Parameter(Mandatory)][string]$Path)
    # LF line ends on every platform, like case-data.json.
    Write-Utf8File $Path (($Audit | ConvertTo-Json -Depth 10).Replace("`r`n", "`n") + "`n")
}

function Approve-DossierReport {
    <#
      Human approval of one report. Verifies that case data, prompt, response and
      draft report still match the hashes in the audit.json of the report folder,
      writes an approved copy of the report without the DRAFT markers, and records
      reviewer, time and new hash. The draft stays.

      Approval is refused while any automated check in audit.json has failed.
      -AcceptFailedChecks with a non-empty -Justification overrides that; the
      override is printed in the approved report and stored in audit.json as a
      documented exception.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
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
    $auditPath = Join-Path $Path 'audit.json'
    if (-not (Test-Path -LiteralPath $auditPath)) {
        $hint = ''
        if (Test-Path -LiteralPath (Join-Path $Path 'run.json')) { $hint = ' This is a run folder; approve each report in its case folder, for example <run folder>/<case id>.' }
        throw "audit.json not found in $Path.$hint"
    }
    $audit = Read-JsonFile $auditPath
    if ($audit.status -ne 'DRAFT') { throw "Report $($audit.caseId) of run $($audit.runId) has status $($audit.status); only DRAFT reports can be approved." }
    if ($audit.PSObject.Properties.Name -notcontains 'checks' -or @($audit.checks).Count -eq 0) {
        throw "audit.json of $($audit.caseId) records no automated checks. Approval refused."
    }
    $failed = @($audit.checks | Where-Object { $_.passed -ne $true })
    $failedText = @($failed | ForEach-Object {
            if ($_.detail) { "$($_.name) ($($_.detail))" } else { [string]$_.name }
        }) -join '; '
    if ($failed.Count -and -not $AcceptFailedChecks) {
        throw ("Approval refused: $($failed.Count) automated check(s) failed: $failedText. The report stays DRAFT. " +
            "Draft it again in a new run, or accept the failures as a documented exception with -AcceptFailedChecks -Justification '<reason>'.")
    }
    if ($AcceptFailedChecks -and -not $failed.Count) {
        throw 'All automated checks passed; -AcceptFailedChecks is not needed. Approval refused.'
    }

    foreach ($name in 'caseData', 'prompt', 'response', 'report') {
        $artifact = $audit.artifacts.$name
        $file = Join-Path $Path $artifact.file
        if (-not (Test-Path -LiteralPath $file)) { throw "Artifact missing: $($artifact.file). Approval refused." }
        $actual = Get-Sha256 $file
        if ($actual -ne $artifact.sha256) { throw "Hash mismatch for $($artifact.file): audit has $($artifact.sha256), file has $actual. The file changed after the run. Approval refused." }
    }

    $approvedAt = Get-UtcStamp
    $m = $script:Markers
    $replace = [ordered]@{
        $m.Banner     = "REVIEWED AND APPROVED by $Reviewer on $approvedAt"
        $m.Status     = 'Status: APPROVED'
        $m.Reviewer   = "Approved by: $Reviewer"
        $m.ApprovedAt = "Approved at (UTC): $approvedAt"
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
        $replace[$m.Banner] = "REVIEWED AND APPROVED WITH A DOCUMENTED EXCEPTION by $Reviewer on $approvedAt"
        $replace[$m.Status] = 'Status: APPROVED with a documented exception (failed automated checks accepted)'
        $replace[$m.Exception] = "Accepted check failures: DOCUMENTED EXCEPTION. Failed checks: $failedText. Justification: $($exception.justification)"
    } else {
        $replace[$m.Exception] = 'Accepted check failures: none (all automated checks passed)'
    }

    $draft = Join-Path $Path $audit.artifacts.report.file
    $approvedName = $audit.artifacts.report.file -replace '-DRAFT\.docx$', '.docx'
    if ($approvedName -eq $audit.artifacts.report.file) { $approvedName = [IO.Path]::GetFileNameWithoutExtension($approvedName) + '-approved.docx' }
    $approved = Join-Path $Path $approvedName
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
        reviewer      = $Reviewer
        approvedAtUtc = $approvedAt
        comment       = $Comment
        osAccount     = [Environment]::UserName
        draftSha256   = $audit.artifacts.report.sha256
        exception     = $exception
    }
    $audit.artifacts | Add-Member -NotePropertyName approvedReport -NotePropertyValue ([pscustomobject]@{ file = $approvedName; sha256 = (Get-Sha256 $approved) }) -Force
    if ($exception) {
        $audit.events += [pscustomobject]@{ atUtc = $approvedAt; event = 'approved-with-exception'; detail = "by $Reviewer; accepted failed checks: $failedText" }
    } else {
        $audit.events += [pscustomobject]@{ atUtc = $approvedAt; event = 'approved'; detail = "by $Reviewer" }
    }
    Save-AuditLog $audit $auditPath
    [pscustomobject]@{ RunId = $audit.runId; CaseId = $audit.caseId; Report = $approved; Sha256 = $audit.artifacts.approvedReport.sha256; ApprovedAt = $approvedAt; Exception = $exception }
}

Export-ModuleMember -Function *-*
