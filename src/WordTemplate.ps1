# Fills a user's own Word template (.docx or .dotx) without Word.
# Dot-sourced by ComplaintDossier.psm1. Works on the OpenXML parts directly:
# word/document.xml, word/header*.xml and word/footer*.xml.
#
# Placeholder rules (see docs/word-template.md):
# - A placeholder is {{Name}}: two curly braces, a name made of letters, digits
#   and underscores that starts with a letter, two closing braces. Spaces inside
#   the braces and upper/lower case do not matter: {{ batch }} = {{Batch}}.
# - Word often splits typed text into several runs (spell check, an edit in the
#   middle, a formatting change). The placeholder is matched on the visible text
#   of the paragraph, so a split placeholder is still found. The inserted value
#   takes the formatting of the first character of the placeholder.
# - A text placeholder (Excel field, date, model ...) can stand anywhere in a
#   paragraph, a table cell, a header or a footer.
# - A block placeholder (AI section, check results ...) that stands alone in its
#   paragraph is replaced by as many paragraphs as the text needs. Each new
#   paragraph keeps the paragraph format of the placeholder paragraph; list lines
#   get the bullet style, headings the heading style named in the configuration.
#   A block placeholder that shares its paragraph with other text is inserted as
#   plain text with line breaks.
# Keep this file ASCII-only: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.

$script:WNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$script:XmlNs = 'http://www.w3.org/XML/1998/namespace'
$script:PlaceholderPattern = '\{\{\s*([A-Za-z][A-Za-z0-9_]*)\s*\}\}'
$script:WordTextParts = '^word/(document|header\d*|footer\d*)\.xml$'

function Get-WordPartsText {
    <# Reads the text parts of a .docx/.dotx: name -> XML text. #>
    param([Parameter(Mandatory)][string]$Path)
    $parts = [ordered]@{}
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -match $script:WordTextParts } | Sort-Object FullName)) {
            $parts[$entry.FullName] = Get-ZipEntryText $zip $entry.FullName
        }
        $parts['styles'] = Get-ZipEntryText $zip 'word/styles.xml'
    } finally { $zip.Dispose() }
    if (-not $parts.Contains('word/document.xml')) { throw "Not a Word document (word/document.xml missing): $([IO.Path]::GetFileName($Path))" }
    $parts
}

function New-WordXmlDocument {
    param([Parameter(Mandatory)][string]$Xml)
    $doc = New-Object Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($Xml)
    $doc
}

function Get-WParagraphTextNodes {
    <# The w:t elements of one paragraph, without those of a text box paragraph nested inside it. #>
    param([Parameter(Mandatory)][Xml.XmlElement]$Paragraph)
    $nodes = New-Object System.Collections.Generic.List[Xml.XmlElement]
    foreach ($t in @($Paragraph.GetElementsByTagName('t', $script:WNs))) {
        $owner = $t.ParentNode
        while ($null -ne $owner -and -not ($owner.LocalName -eq 'p' -and $owner.NamespaceURI -eq $script:WNs)) { $owner = $owner.ParentNode }
        if ([object]::ReferenceEquals($owner, $Paragraph)) { $nodes.Add($t) }
    }
    , $nodes.ToArray()
}

function Get-WordPlaceholders {
    <#
      Lists the placeholders of a Word template: one entry per occurrence with
      Name, Part (document, header1, footer1 ...) and Alone (the placeholder is
      the only text of its paragraph, so it can take a block of paragraphs).
    #>
    param([Parameter(Mandatory)][string]$Path)
    $parts = Get-WordPartsText $Path
    foreach ($name in @($parts.Keys | Where-Object { $_ -ne 'styles' })) {
        $doc = New-WordXmlDocument $parts[$name]
        foreach ($p in @($doc.GetElementsByTagName('p', $script:WNs))) {
            $text = -join @(Get-WParagraphTextNodes $p | ForEach-Object { $_.InnerText })
            if ($text.IndexOf('{{') -lt 0) { continue }
            $found = [regex]::Matches($text, $script:PlaceholderPattern)
            foreach ($m in $found) {
                [pscustomobject]@{
                    Name  = $m.Groups[1].Value
                    Part  = [IO.Path]::GetFileNameWithoutExtension($name)
                    Alone = ($found.Count -eq 1 -and $text.Trim() -eq $m.Value)
                    Text  = $text.Trim()
                }
            }
        }
    }
}

function Get-WordStyleId {
    <#
      Finds a paragraph style in styles.xml by its ID or its name, ignoring case
      ("List Bullet", "ListBullet", "heading 2", "Heading2"). Built-in styles
      keep their English name in the file, also in a German Word. Returns the
      style ID, or '' when the template does not define the style.
    #>
    param([string]$StylesXml, [string]$Style)
    if (-not $StylesXml -or -not $Style) { return '' }
    $doc = New-WordXmlDocument $StylesXml
    $want = ($Style -replace '\s', '').ToLowerInvariant()
    foreach ($s in @($doc.GetElementsByTagName('style', $script:WNs))) {
        if ($s.GetAttribute('type', $script:WNs) -ne 'paragraph') { continue }
        $id = $s.GetAttribute('styleId', $script:WNs)
        $nameNode = @($s.GetElementsByTagName('name', $script:WNs))
        $name = ''
        if ($nameNode.Count) { $name = $nameNode[0].GetAttribute('val', $script:WNs) }
        if (($id -replace '\s', '').ToLowerInvariant() -eq $want -or ($name -replace '\s', '').ToLowerInvariant() -eq $want) { return $id }
    }
    ''
}

function ConvertTo-WordBlockItems {
    <#
      Splits text into paragraph items. Each non-empty line is one paragraph,
      as Copilot answers put one paragraph per line. "## x" is a heading,
      "- x", "* x" and bullet characters are list lines. HTML comments go away.
    #>
    param([AllowEmptyString()][string]$Text)
    $Text = [regex]::Replace($Text, '(?s)<!--.*?-->', '')
    foreach ($line in ($Text -split "`r?`n")) {
        $trim = $line.Trim()
        if (-not $trim) { continue }
        if ($trim -match '^#{1,6}\s+(.+)$') { [pscustomobject]@{ Kind = 'Heading'; Text = $Matches[1].Trim() } }
        elseif ($trim -match ('^[-*' + [char]0x2022 + ']\s+(.+)$')) { [pscustomobject]@{ Kind = 'Bullet'; Text = $Matches[1].Trim() } }
        elseif ($trim -match '^>\s?(.*)$') { [pscustomobject]@{ Kind = 'Text'; Text = $Matches[1].Trim() } }
        else { [pscustomobject]@{ Kind = 'Text'; Text = $trim } }
    }
}

function ConvertTo-PlainBlockText {
    <# Block text for a placeholder that shares its paragraph with other text: lines, no Markdown marks. #>
    param([AllowEmptyString()][string]$Text)
    (@(ConvertTo-WordBlockItems $Text | ForEach-Object {
                $line = $_.Text.Replace('**', '')
                if ($_.Kind -eq 'Bullet') { [string][char]0x2022 + ' ' + $line } else { $line }
            }) -join "`n")
}

function Set-XmlPreserve {
    param([Xml.XmlElement]$Node)
    $attr = $Node.OwnerDocument.CreateAttribute('xml', 'space', $script:XmlNs)
    $attr.Value = 'preserve'
    [void]$Node.Attributes.SetNamedItem($attr)
}

function Set-WTextRange {
    <#
      Replaces the characters Start..Start+Length-1 of a paragraph, whose text is
      spread over the w:t nodes in $Nodes, by $Value. The value goes into the
      node that holds the first character, so it keeps that run's formatting.
    #>
    param([Xml.XmlElement[]]$Nodes, [int]$Start, [int]$Length, [AllowEmptyString()][string]$Value)
    $end = $Start + $Length - 1
    $pos = 0
    $first = -1; $firstOffset = 0; $last = -1; $lastOffset = 0
    for ($i = 0; $i -lt $Nodes.Count; $i++) {
        $len = $Nodes[$i].InnerText.Length
        if ($first -lt 0 -and $Start -lt $pos + $len) { $first = $i; $firstOffset = $Start - $pos }
        if ($first -ge 0 -and $end -lt $pos + $len) { $last = $i; $lastOffset = $end - $pos; break }
        $pos += $len
    }
    if ($first -lt 0 -or $last -lt 0) { throw 'Placeholder position not found in the paragraph text.' }
    if ($first -eq $last) {
        $t = $Nodes[$first].InnerText
        $Nodes[$first].InnerText = $t.Substring(0, $firstOffset) + $Value + $t.Substring($lastOffset + 1)
        Set-XmlPreserve $Nodes[$first]
        return
    }
    $Nodes[$first].InnerText = $Nodes[$first].InnerText.Substring(0, $firstOffset) + $Value
    Set-XmlPreserve $Nodes[$first]
    for ($i = $first + 1; $i -lt $last; $i++) { $Nodes[$i].InnerText = '' }
    $Nodes[$last].InnerText = $Nodes[$last].InnerText.Substring($lastOffset + 1)
    Set-XmlPreserve $Nodes[$last]
}

function Split-WTextLineBreaks {
    <# A w:t holding line breaks becomes w:t, w:br, w:t ... inside the same run. #>
    param([Xml.XmlElement]$Node)
    $parts = $Node.InnerText -split "`r?`n"
    if ($parts.Count -lt 2) { return }
    $doc = $Node.OwnerDocument
    $Node.InnerText = $parts[0]
    $after = $Node
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $br = $doc.CreateElement('w', 'br', $script:WNs)
        [void]$Node.ParentNode.InsertAfter($br, $after)
        $t = $doc.CreateElement('w', 't', $script:WNs)
        Set-XmlPreserve $t
        $t.InnerText = $parts[$i]
        [void]$Node.ParentNode.InsertAfter($t, $br)
        $after = $t
    }
}

function Get-WChild {
    param([Xml.XmlNode]$Parent, [string]$LocalName)
    if ($null -eq $Parent) { return $null }
    foreach ($child in $Parent.ChildNodes) {
        if ($child.LocalName -eq $LocalName -and $child.NamespaceURI -eq $script:WNs) { return $child }
    }
    $null
}

function Set-WParagraphStyle {
    <# Sets pStyle as the first child of pPr (schema order) and drops a direct list numbering. #>
    param([Xml.XmlElement]$PPr, [string]$StyleId)
    $doc = $PPr.OwnerDocument
    foreach ($name in 'pStyle', 'numPr') {
        $old = Get-WChild $PPr $name
        if ($old) { [void]$PPr.RemoveChild($old) }
    }
    $style = $doc.CreateElement('w', 'pStyle', $script:WNs)
    [void]$style.SetAttribute('val', $script:WNs, $StyleId)
    [void]$PPr.PrependChild($style)
}

function New-WRunElement {
    param([Xml.XmlDocument]$Doc, [Xml.XmlElement]$RPr, [string]$Text, [switch]$Bold)
    $run = $Doc.CreateElement('w', 'r', $script:WNs)
    $props = $null
    if ($RPr) { $props = $RPr.CloneNode($true) }
    if ($Bold) {
        if (-not $props) { $props = $Doc.CreateElement('w', 'rPr', $script:WNs) }
        $old = Get-WChild $props 'b'
        if ($old) { [void]$props.RemoveChild($old) }
        $b = $Doc.CreateElement('w', 'b', $script:WNs)
        # Schema order: rStyle, rFonts, b ...
        $anchor = $null
        foreach ($name in 'rFonts', 'rStyle') { if (-not $anchor) { $anchor = Get-WChild $props $name } }
        if ($anchor) { [void]$props.InsertAfter($b, $anchor) } else { [void]$props.PrependChild($b) }
    }
    if ($props) { [void]$run.AppendChild($props) }
    $t = $Doc.CreateElement('w', 't', $script:WNs)
    Set-XmlPreserve $t
    $t.InnerText = $Text
    [void]$run.AppendChild($t)
    $run
}

function Set-WParagraphBlock {
    <#
      Replaces a paragraph that holds only a block placeholder by one paragraph
      per item. Paragraph and run format come from the placeholder paragraph.
    #>
    param([Xml.XmlElement]$Paragraph, [Xml.XmlElement[]]$Nodes, [object[]]$Items, [hashtable]$StyleIds)
    $doc = $Paragraph.OwnerDocument
    $pPr = Get-WChild $Paragraph 'pPr'
    $rPr = $null
    $firstNode = @($Nodes | Where-Object { $_.InnerText.Trim() })[0]
    if ($firstNode) { $rPr = Get-WChild $firstNode.ParentNode 'rPr' }
    $sectPr = $null
    if ($pPr) { $sectPr = Get-WChild $pPr 'sectPr' }
    if (-not @($Items).Count) { $Items = @([pscustomobject]@{ Kind = 'Text'; Text = '' }) }
    $created = New-Object System.Collections.Generic.List[Xml.XmlElement]
    foreach ($item in $Items) {
        $p = $doc.CreateElement('w', 'p', $script:WNs)
        $props = $null
        if ($pPr) {
            $props = $pPr.CloneNode($true)
            $inner = Get-WChild $props 'sectPr'
            if ($inner) { [void]$props.RemoveChild($inner) }
        }
        $styleId = ''
        if ($item.Kind -eq 'Heading') { $styleId = [string]$StyleIds.Heading }
        if ($item.Kind -eq 'Bullet') { $styleId = [string]$StyleIds.Bullet }
        if ($styleId) {
            if (-not $props) { $props = $doc.CreateElement('w', 'pPr', $script:WNs) }
            Set-WParagraphStyle $props $styleId
        }
        if ($props) { [void]$p.AppendChild($props) }
        $text = $item.Text
        if ($item.Kind -eq 'Bullet' -and -not $styleId) { $text = [string][char]0x2022 + ' ' + $text }
        # A heading with its own style takes the style's font, not the body text format.
        $runProps = $rPr
        if ($item.Kind -eq 'Heading' -and $styleId) { $runProps = $null }
        $boldAll = ($item.Kind -eq 'Heading' -and -not $styleId)
        foreach ($part in [regex]::Split($text, '(\*\*[^*]+\*\*)')) {
            if (-not $part) { continue }
            if ($part -match '^\*\*(.+)\*\*$') { [void]$p.AppendChild((New-WRunElement $doc $runProps $Matches[1] -Bold)) }
            else { [void]$p.AppendChild((New-WRunElement $doc $runProps $part -Bold:$boldAll)) }
        }
        [void]$Paragraph.ParentNode.InsertBefore($p, $Paragraph)
        $created.Add($p)
    }
    if ($sectPr) {
        # A section break stays on the last paragraph that replaces the placeholder.
        $last = $created[$created.Count - 1]
        $props = Get-WChild $last 'pPr'
        if (-not $props) { $props = $doc.CreateElement('w', 'pPr', $script:WNs); [void]$last.PrependChild($props) }
        [void]$props.AppendChild($sectPr.CloneNode($true))
    }
    [void]$Paragraph.ParentNode.RemoveChild($Paragraph)
}

function Update-WordPartXml {
    <#
      Replaces the placeholders in one part. $Values: name -> text (inline);
      $Blocks: name -> text (paragraphs). Names are matched without regard to
      case. Returns the new XML and the names that were filled or unknown.
    #>
    param([string]$Xml, [hashtable]$Values, [hashtable]$Blocks, [hashtable]$StyleIds)
    $doc = New-WordXmlDocument $Xml
    $filled = New-Object System.Collections.Generic.List[string]
    $unknown = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($doc.GetElementsByTagName('p', $script:WNs))) {
        if ($null -eq $p.ParentNode) { continue }
        $nodes = Get-WParagraphTextNodes $p
        if (-not $nodes.Count) { continue }
        $text = -join @($nodes | ForEach-Object { $_.InnerText })
        if ($text.IndexOf('{{') -lt 0) { continue }
        $found = [regex]::Matches($text, $script:PlaceholderPattern)
        if (-not $found.Count) { continue }
        $name = $found[0].Groups[1].Value
        if ($found.Count -eq 1 -and $text.Trim() -eq $found[0].Value -and $Blocks.ContainsKey($name)) {
            Set-WParagraphBlock -Paragraph $p -Nodes $nodes -Items @(ConvertTo-WordBlockItems ([string]$Blocks[$name])) -StyleIds $StyleIds
            $filled.Add($name)
            continue
        }
        # Right to left, so the positions of earlier placeholders stay valid.
        for ($i = $found.Count - 1; $i -ge 0; $i--) {
            $name = $found[$i].Groups[1].Value
            if ($Values.ContainsKey($name)) { $value = [string]$Values[$name] }
            elseif ($Blocks.ContainsKey($name)) { $value = ConvertTo-PlainBlockText ([string]$Blocks[$name]) }
            else { $unknown.Add($name); continue }
            Set-WTextRange -Nodes $nodes -Start $found[$i].Index -Length $found[$i].Length -Value $value
            $filled.Add($name)
        }
        foreach ($n in $nodes) { if ($n.InnerText.IndexOf("`n") -ge 0) { Split-WTextLineBreaks $n } }
    }
    [pscustomobject]@{ Xml = $doc.OuterXml; Filled = $filled.ToArray(); Unknown = $unknown.ToArray() }
}

function Add-WordBodyParagraphs {
    <#
      Adds paragraphs to the body of document.xml: at the top (-AtStart) or
      before the final section properties. Used when a template has no
      {{DraftNotice}} or no {{CheckSummary}}, so a report is never without them.
    #>
    param([string]$Xml, [object[]]$Items, [switch]$AtStart, [string]$Color, [switch]$Bold, [hashtable]$StyleIds)
    $doc = New-WordXmlDocument $Xml
    $body = @($doc.GetElementsByTagName('body', $script:WNs))[0]
    $before = $null
    if ($AtStart) { $before = $body.FirstChild } else { $before = Get-WChild $body 'sectPr' }
    $rPr = $null
    if ($Color) {
        $rPr = $doc.CreateElement('w', 'rPr', $script:WNs)
        $c = $doc.CreateElement('w', 'color', $script:WNs)
        [void]$c.SetAttribute('val', $script:WNs, $Color)
        [void]$rPr.AppendChild($c)
    }
    foreach ($item in $Items) {
        $p = $doc.CreateElement('w', 'p', $script:WNs)
        $styleId = ''
        if ($item.Kind -eq 'Heading') { $styleId = [string]$StyleIds.Heading }
        if ($item.Kind -eq 'Bullet') { $styleId = [string]$StyleIds.Bullet }
        if ($styleId) {
            $props = $doc.CreateElement('w', 'pPr', $script:WNs)
            Set-WParagraphStyle $props $styleId
            [void]$p.AppendChild($props)
        }
        $text = $item.Text
        if ($item.Kind -eq 'Bullet' -and -not $styleId) { $text = [string][char]0x2022 + ' ' + $text }
        [void]$p.AppendChild((New-WRunElement $doc $rPr $text -Bold:($Bold -or ($item.Kind -eq 'Heading' -and -not $styleId))))
        if ($before) { [void]$body.InsertBefore($p, $before) } else { [void]$body.AppendChild($p) }
    }
    $doc.OuterXml
}

function Set-CoreContentStatus {
    <# Sets the document status property (File > Info > Status in Word), when the file has core properties. #>
    param([string]$Xml, [string]$Status)
    $value = ConvertTo-XmlText $Status
    if ($Xml -match '<cp:contentStatus\s*/>|<cp:contentStatus>.*?</cp:contentStatus>') {
        return [regex]::Replace($Xml, '<cp:contentStatus\s*/>|<cp:contentStatus>.*?</cp:contentStatus>', "<cp:contentStatus>$value</cp:contentStatus>")
    }
    if ($Xml -match 'xmlns:cp=') { return $Xml.Replace('</cp:coreProperties>', "<cp:contentStatus>$value</cp:contentStatus></cp:coreProperties>") }
    $Xml
}

function Invoke-WordTemplate {
    <#
      Writes $OutputPath from the template $TemplatePath with all placeholders
      filled. Fails closed: an unknown placeholder, a placeholder left in the
      result or a template without a body stops the report.
      $Values: inline text per name. $Blocks: paragraph text per name.
      $Notice: text put at the top of the body when the template has no
      {{DraftNotice}}. $Appendix: text added at the end of the body when the
      template has no {{CheckSummary}}.
    #>
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][hashtable]$Values,
        [Parameter(Mandatory)][hashtable]$Blocks,
        [hashtable]$Styles = @{},
        [string]$Notice = '',
        [string]$Appendix = '',
        [string]$ContentStatus = ''
    )
    $parts = Get-WordPartsText $TemplatePath
    $names = @(Get-WordPlaceholders $TemplatePath | ForEach-Object { $_.Name.ToLowerInvariant() })
    $styleIds = @{
        Heading = (Get-WordStyleId $parts['styles'] ([string]$Styles['Heading']))
        Bullet  = (Get-WordStyleId $parts['styles'] ([string]$Styles['Bullet']))
    }
    $unknown = New-Object System.Collections.Generic.List[string]
    $script:wordTemplateUnknown = $unknown
    Update-ZipEntries -Source $TemplatePath -Destination $OutputPath -Transform {
        param($name, $text)
        if ($name -eq '[Content_Types].xml') {
            # A .dotx template becomes a normal document.
            return $text.Replace('wordprocessingml.template.main+xml', 'wordprocessingml.document.main+xml')
        }
        if ($name -eq 'docProps/core.xml' -and $ContentStatus) { return Set-CoreContentStatus $text $ContentStatus }
        if ($name -notmatch $script:WordTextParts) { return $text }
        $result = Update-WordPartXml -Xml $text -Values $Values -Blocks $Blocks -StyleIds $styleIds
        foreach ($u in $result.Unknown) { $script:wordTemplateUnknown.Add($u) }
        $xml = $result.Xml
        if ($name -eq 'word/document.xml') {
            if ($Notice -and $names -notcontains 'draftnotice') {
                $xml = Add-WordBodyParagraphs -Xml $xml -Items @([pscustomobject]@{ Kind = 'Text'; Text = $Notice }) -AtStart -Color 'C00000' -Bold -StyleIds $styleIds
            }
            if ($Appendix -and $names -notcontains 'checksummary') {
                $xml = Add-WordBodyParagraphs -Xml $xml -Items @(ConvertTo-WordBlockItems $Appendix) -StyleIds $styleIds
            }
        }
        $xml
    }
    if ($unknown.Count) {
        Remove-Item -LiteralPath $OutputPath -Force
        throw "Unknown placeholder(s) in the Word template $([IO.Path]::GetFileName($TemplatePath)): $((@($unknown | Sort-Object -Unique | ForEach-Object { '{{' + $_ + '}}' })) -join ', '). Run Setup to see which names exist."
    }
    $left = [regex]::Matches((Get-DocxText $OutputPath), $script:PlaceholderPattern)
    if ($left.Count) {
        Remove-Item -LiteralPath $OutputPath -Force
        throw "Placeholder left in the report after filling: $($left[0].Value)"
    }
}
