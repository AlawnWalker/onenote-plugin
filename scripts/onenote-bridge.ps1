# =============================================================================
# dsh-onenote - OneNote COM bridge
# -----------------------------------------------------------------------------
# Single entry point used by the dsh-onenote DSH plugin. It receives a JSON
# request file, talks to the local OneNote desktop application through COM, and
# writes a JSON response file. Keep this file pure ASCII: Chinese text always
# travels inside the JSON payloads (PowerShell 5.1 needs a BOM for source files
# that contain non-ASCII literals).
# =============================================================================
param(
    [Parameter(Mandatory = $true)][string]$Request,
    [Parameter(Mandatory = $true)][string]$Response
)

$ErrorActionPreference = 'Stop'
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:NS = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$script:App = $null

function Fail([string]$code, [string]$message) {
    throw ('__DSH__' + $code + '|' + $message)
}

function Read-JsonFile([string]$path) {
    $text = [IO.File]::ReadAllText($path, $script:Utf8)
    return ($text | ConvertFrom-Json)
}

function Write-JsonFile([string]$path, $obj) {
    $json = $obj | ConvertTo-Json -Depth 20 -Compress
    [IO.File]::WriteAllText($path, $json, $script:Utf8)
}

function Get-App {
    if ($null -eq $script:App) {
        $script:App = New-Object -ComObject OneNote.Application
    }
    return $script:App
}

# OneNote HRESULT codes (2013 object model) -> stable, human-readable codes.
$script:HResultMap = @{
    '80042000' = @{ code = 'MALFORMED_XML'; hint = 'the OneNote XML is not well-formed' }
    '80042001' = @{ code = 'INVALID_XML'; hint = 'the OneNote XML is invalid' }
    '80042002' = @{ code = 'CREATE_SECTION_FAILED'; hint = 'the section could not be created' }
    '80042003' = @{ code = 'OPEN_SECTION_FAILED'; hint = 'the section could not be opened' }
    '80042004' = @{ code = 'SECTION_NOT_FOUND'; hint = 'the section does not exist' }
    '80042005' = @{ code = 'PAGE_NOT_FOUND'; hint = 'the page does not exist' }
    '80042006' = @{ code = 'FILE_NOT_FOUND'; hint = 'the file does not exist (check the path)' }
    '8004200B' = @{ code = 'SECTION_READ_ONLY'; hint = 'the section is read-only (typical for standalone/network sections)' }
    '8004200C' = @{ code = 'PAGE_READ_ONLY'; hint = 'the page is read-only' }
    '8004200E' = @{ code = 'PAGE_OBJECT_NOT_FOUND'; hint = 'the page object does not exist' }
    '80042010' = @{ code = 'MODIFIED_CONFLICT'; hint = 'the page was modified by someone else' }
    '80042014' = @{ code = 'OBJECT_NOT_FOUND'; hint = 'the object does not exist' }
    '80042015' = @{ code = 'NOTEBOOK_NOT_FOUND'; hint = 'the notebook does not exist' }
    '80042018' = @{ code = 'GROUP_NOT_FOUND'; hint = 'the section group/folder does not exist' }
    '80042019' = @{ code = 'INVALID_QUERY'; hint = 'the query is invalid' }
    '8004201A' = @{ code = 'FILE_EXISTS'; hint = 'the target file already exists' }
    '8004201B' = @{ code = 'SECTION_LOCKED'; hint = 'the section is encrypted and locked' }
    '8004201C' = @{ code = 'DISABLED_BY_POLICY'; hint = 'the action is disabled by policy' }
    '80042020' = @{ code = 'INVALID_XML_SCHEMA'; hint = 'the OneNote XML schema is invalid' }
    '80042023' = @{ code = 'TIMEOUT'; hint = 'OneNote timed out' }
    '80042030' = @{ code = 'MODAL_DIALOG'; hint = 'a modal dialog is blocking OneNote' }
    '800706BE' = @{ code = 'RPC_CALL_FAILED'; hint = 'the remote procedure call failed (OneNote may be busy or restarting)' }
    '80010001' = @{ code = 'RPC_CALL_REJECTED'; hint = 'the call was rejected (OneNote is busy)' }
    '80080005' = @{ code = 'COM_SERVER_EXEC_FAILURE'; hint = 'the OneNote COM server could not start (blocked by a sandbox?)' }
}

# PowerShell wraps a COM failure in a MethodInvocationException whose HResult is
# the generic CLR code; the OneNote HRESULT lives on an inner exception. Walk the
# chain and prefer the first code that the map knows.
function Get-DeepHResult($ex) {
    $first = '00000000'
    $cur = $ex
    while ($null -ne $cur) {
        if ($null -ne $cur.HResult) {
            try {
                # HRESULT is a signed Int32; [uint32] on a negative value throws,
                # so widen to Int64 and add 2^32 before formatting as hex.
                $hv = [int64]$cur.HResult
                if ($hv -lt 0) { $hv += 4294967296 }
                $hex = '{0:X8}' -f $hv
                if ($script:HResultMap.ContainsKey($hex)) { return $hex }
                if ($first -eq '00000000' -and $hex -ne '80131501') { $first = $hex }
            } catch { }
        }
        $cur = $cur.InnerException
    }
    return $first
}

function Get-ComError($ex) {
    $hex = Get-DeepHResult $ex
    $mapped = $script:HResultMap[$hex]
    if ($null -ne $mapped) {
        return [ordered]@{ code = $mapped.code; message = ($mapped.hint + ': ' + $ex.Message); hresult = '0x' + $hex }
    }
    return [ordered]@{ code = 'ONENOTE_ERROR'; message = $ex.Message; hresult = '0x' + $hex }
}

function New-XmlDoc([string]$xml) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    try {
        $doc.LoadXml($xml)
    } catch {
        Fail 'BAD_XML' $_.Exception.Message
    }
    # XmlDocument is IEnumerable; the leading comma stops PowerShell from
    # unrolling it into its child nodes on return.
    return ,$doc
}

function New-NsMgr($doc) {
    $nm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nm.AddNamespace('one', $script:NS)
    # XmlNamespaceManager is IEnumerable as well - return it as one object.
    return ,$nm
}

function Get-HierarchyXml([int]$scope, [string]$startId) {
    $xml = ''
    (Get-App).GetHierarchy($startId, $scope, [ref]$xml)
    return $xml
}

# OneNote stores outline text as HTML-ish content: it entity-encodes < > & "
# and wraps runs that need another font in <span style=...>. Decode both, or the
# text read back differs from what the user sees in the notebook.
function Convert-OneNoteText([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return '' }
    $t = $text
    $t = [regex]::Replace($t, '</?(span|b|strong|i|em|u|sup|sub|a|font)(\s[^>]*)?>', '')
    $t = [regex]::Replace($t, '<br\s*/?>', "`n")
    $t = $t.Replace('&lt;', '<').Replace('&gt;', '>').Replace('&quot;', '"').Replace('&apos;', "'").Replace('&#39;', "'").Replace('&nbsp;', ' ')
    $t = [regex]::Replace($t, '&#x([0-9a-fA-F]+);', {
            param($m)
            $cp = [Convert]::ToInt32($m.Groups[1].Value, 16)
            if ($cp -gt 65535) { [char]::ConvertFromUtf32($cp) } else { [char]$cp }
        })
    $t = [regex]::Replace($t, '&#(\d+);', {
            param($m)
            $cp = [int]$m.Groups[1].Value
            if ($cp -gt 65535) { [char]::ConvertFromUtf32($cp) } else { [char]$cp }
        })
    $t = $t.Replace('&amp;', '&')
    return $t
}

function Get-TextOf($node) {
    $parts = @()
    foreach ($t in $node.SelectNodes('.//*[local-name()="T"]')) {
        $parts += (Convert-OneNoteText $t.InnerText)
    }
    return ($parts -join '')
}

function Get-DirectText($oe) {
    $parts = @()
    foreach ($t in $oe.SelectNodes('./*[local-name()="T"]')) {
        $parts += (Convert-OneNoteText $t.InnerText)
    }
    return ($parts -join '')
}

# Effective font size of one outline element: its own style wins, then the
# style of its direct text runs (OneNote stores run formatting there).
function Get-OeFontSize($oe) {
    $styles = New-Object System.Collections.ArrayList
    if ($oe.HasAttribute('style')) { [void]$styles.Add($oe.GetAttribute('style')) }
    foreach ($t in $oe.SelectNodes('./*[local-name()="T"]')) {
        if ($t.HasAttribute('style')) { [void]$styles.Add($t.GetAttribute('style')) }
    }
    foreach ($s in $styles) {
        $m = [regex]::Match([string]$s, 'font-size:\s*([0-9]+(?:\.[0-9]+)?)\s*pt')
        if ($m.Success) { return [double]$m.Groups[1].Value }
    }
    return 0
}

# ---------------------------------------------------------------------------
# XML escaping helpers
# ---------------------------------------------------------------------------

function ConvertTo-CData([string]$text) {
    if ($null -eq $text) { $text = '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $text.ToCharArray()) {
        $code = [int]$ch
        # XML 1.0 forbids most control characters; drop them instead of failing.
        if ($code -eq 9 -or $code -eq 10 -or $code -eq 13 -or $code -ge 32) {
            [void]$sb.Append($ch)
        }
    }
    $clean = $sb.ToString().Replace(']]>', ']]]]><![CDATA[>')
    return '<![CDATA[' + $clean + ']]>'
}

function ConvertTo-Attr([string]$value) {
    if ($null -eq $value) { return '' }
    return $value.Replace('&', '&amp;').Replace('"', '&quot;').Replace('<', '&lt;').Replace('>', '&gt;')
}

# ---------------------------------------------------------------------------
# Page XML construction from the plugin's block model
# ---------------------------------------------------------------------------

# Heading paragraphs are real formatted paragraphs: font size plus paragraph
# spacing. Only the ASCII style string is used (no font-family) so this script
# stays ASCII-only; OneNote merges the size into the page's default font.
# Sizes are RELATIVE to the body size (see Resolve-BaseFontSize): h3 = base + 2,
# h2 = base + 4, so a 12pt notebook gets 14/16 while a 10pt one gets 12/14.
$script:HeadingSpace = @{
    h2 = @{ before = '8.0'; after = '5.0' }
    h3 = @{ before = '5.0'; after = '5.0' }
}
$script:HeadingDelta = @{ h2 = 4; h3 = 2 }
$script:BaseFontSize = 12.0

function Format-PtSize([double]$value) {
    return ('{0:0.0}pt' -f $value)
}

# OneNote's global default body font size lives in the registry
# (Options > Advanced > default font). Returns 0 when it cannot be read.
function Get-OneNoteDefaultFontSize {
    try {
        $items = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\OneNote\Options\Editing' -Name DefaultFontSize -ErrorAction SilentlyContinue
    } catch {
        return 0
    }
    if ($null -eq $items) { return 0 }
    foreach ($item in @($items)) {
        $value = 0
        try { $value = [double]$item.DefaultFontSize } catch { $value = 0 }
        if ($value -gt 0) { return $value }
    }
    return 0
}

# Body size of an already-populated page: its 'p' quick style, else the size
# most paragraphs actually use. Returns 0 when the page has no body text yet.
function Get-DocBaseFontSize($doc) {
    $nm = New-NsMgr $doc
    foreach ($def in $doc.SelectNodes('//one:QuickStyleDef', $nm)) {
        if ($def.GetAttribute('name') -eq 'p') {
            $fs = $def.GetAttribute('fontSize')
            if ($fs -ne '') {
                try { return [double]$fs } catch { }
            }
        }
    }
    $counts = @{}
    # Only body outlines count - the Title outline carries the 17-20pt title run.
    foreach ($oe in $doc.SelectNodes('//one:Page/one:Outline//one:OE', $nm)) {
        if ($null -eq $oe.SelectSingleNode('./*[local-name()="T"]', $nm)) { continue }
        if ($null -ne $oe.SelectSingleNode('./*[local-name()="List"]', $nm)) { continue }
        $size = Get-OeFontSize $oe
        if ($size -le 0) { continue }
        if ($counts.ContainsKey($size)) { $counts[$size] = $counts[$size] + 1 } else { $counts[$size] = 1 }
    }
    $bestSize = 0
    $bestCount = 0
    foreach ($key in $counts.Keys) {
        if ($counts[$key] -gt $bestCount) {
            $bestCount = $counts[$key]
            $bestSize = $key
        }
    }
    return $bestSize
}

# Body size used as the heading baseline when writing to $pageId: the page's own
# body style on an existing page, otherwise a sample of pages already in the same
# section (that is the size OneNote actually renders with), and only then the
# configured default from the registry. OneNote caches its default font at
# startup, so the registry alone can be stale.
function Resolve-BaseFontSize([string]$pageId, [string]$sectionId) {
    if ($pageId -ne '') {
        try {
            $xml = ''
            (Get-App).GetPageContent($pageId, [ref]$xml, 0)
            if ($xml -ne '') {
                $size = Get-DocBaseFontSize (New-XmlDoc $xml)
                if ($size -gt 0) { return $size }
            }
        } catch { }
    }
    if ($sectionId -ne '') {
        try {
            $app = Get-App
            $sectionXml = ''
            $app.GetHierarchy($sectionId, 4, [ref]$sectionXml)
            if ($sectionXml -ne '') {
                $sectionDoc = New-XmlDoc $sectionXml
                $sectionNm = New-NsMgr $sectionDoc
                $sampled = 0
                foreach ($pageNode in $sectionDoc.SelectNodes('//one:Page', $sectionNm)) {
                    if ($sampled -ge 5) { break }
                    $sampled++
                    $candidate = [string]$pageNode.GetAttribute('ID')
                    if ($candidate -eq $pageId) { continue }
                    $pageXml = ''
                    try { $app.GetPageContent($candidate, [ref]$pageXml, 0) } catch { continue }
                    if ($pageXml -eq '') { continue }
                    $size = Get-DocBaseFontSize (New-XmlDoc $pageXml)
                    if ($size -gt 0) { return $size }
                }
            }
        } catch { }
    }
    $default = Get-OneNoteDefaultFontSize
    if ($default -gt 0) { return $default }
    return 12.0
}

function Render-TitleXml([string]$title) {
    return '<one:Title><one:OE><one:T>' + (ConvertTo-CData $title) + '</one:T></one:OE></one:Title>'
}

function Convert-RowsToTableXml($rows) {
    $rows = @($rows)
    if ($rows.Count -eq 0) { return '' }
    $ncols = 0
    foreach ($row in $rows) {
        $n = @($row).Count
        if ($n -gt $ncols) { $ncols = $n }
    }
    if ($ncols -eq 0) { return '' }
    $width = [Math]::Round(680.0 / $ncols, 1)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<one:Table bordersVisible="true" hasHeaderRow="true"><one:Columns>')
    for ($i = 0; $i -lt $ncols; $i++) {
        [void]$sb.Append('<one:Column index="' + $i + '" width="' + $width + '"/>')
    }
    [void]$sb.Append('</one:Columns>')
    foreach ($row in $rows) {
        $cells = @($row)
        [void]$sb.Append('<one:Row>')
        for ($i = 0; $i -lt $ncols; $i++) {
            $cell = ''
            if ($i -lt $cells.Count -and $null -ne $cells[$i]) { $cell = [string]$cells[$i] }
            # OneNote entity-encodes ">" inside table cells and the entity shows up
            # verbatim when the page is read back, so write the full-width form.
            $cell = $cell.Replace('>', [string][char]0xFF1E)
            [void]$sb.Append('<one:Cell><one:OEChildren><one:OE><one:T>' + (ConvertTo-CData $cell) + '</one:T></one:OE></one:OEChildren></one:Cell>')
        }
        [void]$sb.Append('</one:Row>')
    }
    [void]$sb.Append('</one:Table>')
    return $sb.ToString()
}

function Get-BulletValueForLevel([int]$depth) {
    if ($null -ne $script:BulletValues -and $script:BulletValues.Count -gt 0) {
        $index = $depth
        if ($index -ge $script:BulletValues.Count) { $index = $script:BulletValues.Count - 1 }
        if ($index -lt 0) { $index = 0 }
        return [string]$script:BulletValues[$index]
    }
    return [string]$script:BulletValue
}

function Render-OENodes($nodes, $sb, [int]$depth = 0) {
    foreach ($node in $nodes) {
        $type = [string]$node.type
        if ($type -eq 'table') {
            $tableXml = Convert-RowsToTableXml $node.rows
            if ($tableXml -ne '') {
                [void]$sb.Append('<one:OE>' + $tableXml + '</one:OE>')
            }
            continue
        }
        $text = [string]$node.text
        $listXml = ''
        # Blocks carry the level they had in the markdown source; fall back to the
        # render depth when a node has none.
        $bulletDepth = $depth
        if ($null -ne $node.level) { $bulletDepth = [int]$node.level }
        $levelBullet = Get-BulletValueForLevel $bulletDepth
        if ($type -eq 'bullet') {
            $listXml = '<one:List><one:Bullet bullet="' + $levelBullet + '" fontSize="11.0"/></one:List>'
        } elseif ($type -eq 'number') {
            $listXml = '<one:List><one:Bullet bullet="' + $levelBullet + '" fontSize="11.0"/></one:List>'
            $ordinal = 1
            if ($null -ne $node.ordinal -and "$($node.ordinal)" -ne '') { $ordinal = [int]$node.ordinal }
            $text = "$ordinal. " + $text
        }
        $headingXml = ''
        if ($script:HeadingDelta.ContainsKey($type)) {
            $delta = [int]$script:HeadingDelta[$type]
            $space = $script:HeadingSpace[$type]
            $headingXml = ' style="font-size:' + (Format-PtSize ($script:BaseFontSize + $delta)) + '" spaceBefore="' + $space.before + '" spaceAfter="' + $space.after + '"'
        }
        [void]$sb.Append('<one:OE' + $headingXml + '>' + $listXml + '<one:T>' + (ConvertTo-CData $text) + '</one:T>')
        $children = @($node.children)
        if ($null -ne $node.children -and $children.Count -gt 0) {
            $inner = New-Object System.Text.StringBuilder
            Render-OENodes $children $inner ($depth + 1)
            [void]$sb.Append('<one:OEChildren>' + $inner.ToString() + '</one:OEChildren>')
        }
        [void]$sb.Append('</one:OE>')
    }
}

function Convert-BlocksToOutline([object]$blocks) {
    $blocks = @($blocks)
    if ($blocks.Count -eq 0) { return '' }
    $root = New-Object System.Collections.ArrayList
    # containers[L] receives the nodes of nesting level L; containers[0] is the
    # current section. A heading opens a new section when NestUnderHeadings is on.
    $containers = New-Object System.Collections.ArrayList
    [void]$containers.Add($root)
    $headingOpen = $false
    foreach ($block in $blocks) {
        $type = [string]$block.type
        $node = @{
            type     = $type
            text     = [string]$block.text
            ordinal  = $block.ordinal
            rows     = $block.rows
            level    = 0
            children = (New-Object System.Collections.ArrayList)
        }
        if ($null -ne $block.level) { $node.level = [int]$block.level }
        if ($node.level -lt 0) { $node.level = 0 }

        if ($type -eq 'h2' -or $type -eq 'h3') {
            [void]$root.Add($node)
            while ($containers.Count -gt 1) { $containers.RemoveAt($containers.Count - 1) }
            if ($script:NestUnderHeadings) {
                [void]$containers.Add($node.children)
                $headingOpen = $true
            }
            continue
        }

        $level = $node.level
        # Content under an open heading lives one level deeper than its markdown
        # level, so the whole section body is indented by OneNote.
        if ($headingOpen) { $level += 1 }
        # A paragraph may sit deeper than the current list nesting; clamp instead
        # of creating a stray level.
        if ($level -gt $containers.Count - 1) { $level = $containers.Count - 1 }
        while ($containers.Count - 1 -gt $level) { $containers.RemoveAt($containers.Count - 1) }
        [void]$containers[$level].Add($node)
        [void]$containers.Add($node.children)
    }
    $sb = New-Object System.Text.StringBuilder
    Render-OENodes $root $sb
    if ($sb.Length -eq 0) { return '' }
    return '<one:Outline><one:OEChildren>' + $sb.ToString() + '</one:OEChildren></one:Outline>'
}

# Builds <one:PageSettings>. The schema only accepts either <Automatic/> or the
# Orientation + Dimensions + Margins sequence, so a size must carry all three.
function Render-PageSettingsXml([string]$pageColor, $pageSize) {
    $colorAttr = ''
    if ($null -ne $pageColor -and $pageColor -ne '') {
        $colorAttr = ' color="' + (ConvertTo-Attr $pageColor) + '"'
    }
    $children = ''
    if ($null -ne $pageSize) {
        if ($pageSize -is [string] -and ([string]$pageSize).ToLower() -eq 'automatic') {
            # Explicitly go back to OneNote's Automatic page size.
            $children += '<one:PageSize><one:Automatic/></one:PageSize>'
        } else {
            $w = 0
            $h = 0
            if ($null -ne $pageSize.width) { try { $w = [double]$pageSize.width } catch { } }
            if ($null -ne $pageSize.height) { try { $h = [double]$pageSize.height } catch { } }
            if ($w -gt 0 -and $h -gt 0) {
                $landscape = 'false'
                if ($null -ne $pageSize.landscape -and [bool]$pageSize.landscape) { $landscape = 'true' }
                $children += '<one:PageSize><one:Orientation landscape="' + $landscape + '"/><one:Dimensions width="' + $w + '" height="' + $h + '"/>'
                if ($null -ne $pageSize.margins) {
                    $m = $pageSize.margins
                    $top = 36; $bottom = 36; $left = 72; $right = 72
                    if ($null -ne $m.top) { try { $top = [double]$m.top } catch { } }
                    if ($null -ne $m.bottom) { try { $bottom = [double]$m.bottom } catch { } }
                    if ($null -ne $m.left) { try { $left = [double]$m.left } catch { } }
                    if ($null -ne $m.right) { try { $right = [double]$m.right } catch { } }
                    $children += '<one:Margins top="' + $top + '" bottom="' + $bottom + '" left="' + $left + '" right="' + $right + '"/>'
                }
                $children += '</one:PageSize>'
            }
        }
    }
    if ($children -ne '') {
        return '<one:PageSettings RTL="false"' + $colorAttr + '>' + $children + '<one:RuleLines visible="false"/></one:PageSettings>'
    }
    if ($colorAttr -ne '') {
        # Only the color is specified; OneNote keeps every other page property.
        return '<one:PageSettings RTL="false"' + $colorAttr + '/>'
    }
    return ''
}

function New-PagePatchXml([string]$pageId, [string]$title, [string]$outlineXml, [string]$settingsXml) {
    $titleXml = ''
    if ($null -ne $title -and $title -ne '') { $titleXml = Render-TitleXml $title }
    if ($null -eq $settingsXml) { $settingsXml = '' }
    if ($titleXml -eq '' -and ($null -eq $outlineXml -or $outlineXml -eq '') -and $settingsXml -eq '') { return '' }
    return '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '">' + $settingsXml + $titleXml + $outlineXml + '</one:Page>'
}

# Large pages are written in batches so one oversized XML string cannot fail the
# whole page. Each batch becomes its own one:Outline (title only on the first).
$script:WriteChunkSize = 150

function Convert-BlocksToOutlineXmls([object]$blocks) {
    $blocks = @($blocks)
    $outlines = @()
    # Split only at top-level blocks (level 0) so a bullet's nested children
    # never get separated from their parent by a chunk boundary.
    $current = New-Object System.Collections.ArrayList
    foreach ($block in $blocks) {
        $type = [string]$block.type
        $level = 0
        if ($type -eq 'bullet' -or $type -eq 'number') {
            if ($null -ne $block.level) { $level = [int]$block.level }
        }
        if ($current.Count -ge $script:WriteChunkSize -and $level -eq 0) {
            $outline = Convert-BlocksToOutline $current
            if ($outline -ne '') { $outlines += $outline }
            $current = New-Object System.Collections.ArrayList
        }
        [void]$current.Add($block)
    }
    $outline = Convert-BlocksToOutline $current
    if ($outline -ne '') { $outlines += $outline }
    return ,$outlines
}

function Invoke-WritePageContent([string]$pageId, [string]$title, [object]$blocks, [string]$pageColor, $pageSize, [string]$sectionId) {
    $app = Get-App
    $settings = Render-PageSettingsXml $pageColor $pageSize
    # Headings are sized relative to the body: resolve the baseline before the
    # outlines are rendered (h3 = base + 2, h2 = base + 4).
    if ($null -eq $sectionId) { $sectionId = '' }
    $script:BaseFontSize = Resolve-BaseFontSize $pageId $sectionId
    $outlines = Convert-BlocksToOutlineXmls $blocks
    if ($outlines.Count -eq 0) {
        if (($null -ne $title -and $title -ne '') -or $settings -ne '') {
            $patch = New-PagePatchXml $pageId $title '' $settings
            if ($patch -ne '') { $app.UpdatePageContent($patch, 0) }
        }
        return 0
    }
    for ($i = 0; $i -lt $outlines.Count; $i++) {
        $chunkTitle = ''
        if ($i -eq 0) { $chunkTitle = $title }
        $chunkSettings = ''
        if ($i -eq 0) { $chunkSettings = $settings }
        $patch = New-PagePatchXml $pageId $chunkTitle $outlines[$i] $chunkSettings
        if ($patch -ne '') { $app.UpdatePageContent($patch, 0) }
    }
    return $outlines.Count
}

# ---------------------------------------------------------------------------
# Page XML -> block model
# ---------------------------------------------------------------------------

function Convert-TableToBlock($table) {
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $table.SelectNodes('./*[local-name()="Row"]')) {
        $cells = New-Object System.Collections.ArrayList
        foreach ($cell in $row.SelectNodes('./*[local-name()="Cell"]')) {
            $texts = @()
            foreach ($t in $cell.SelectNodes('.//*[local-name()="T"]')) { $texts += $t.InnerText }
            [void]$cells.Add(($texts -join ' '))
        }
        [void]$rows.Add(@($cells))
    }
    return ([ordered]@{ type = 'table'; rows = @($rows); level = 0 })
}

function Convert-ContainerToBlocks($container, [int]$level) {
    $out = New-Object System.Collections.ArrayList
    foreach ($child in $container.ChildNodes) {
        if ($child.NodeType -ne 'Element') { continue }
        $localName = $child.LocalName
        if ($localName -eq 'OE') {
            $text = Get-DirectText $child
            $oid = $child.GetAttribute('objectID')
            $kind = 'para'
            $list = $child.SelectSingleNode('./*[local-name()="List"]')
            if ($null -ne $list) {
                if ($null -ne $list.SelectSingleNode('./*[local-name()="Bullet"]')) { $kind = 'bullet' }
                elseif ($null -ne $list.SelectSingleNode('./*[local-name()="Number"]')) { $kind = 'number' }
            }
            if ($kind -eq 'para' -and $text -ne '') {
                # Formatted headings written by this plugin sit above the body
                # size: h3 = base + 2, h2 = base + 4.
                $size = Get-OeFontSize $child
                if ($size -ge ($script:BaseFontSize + 3)) { $kind = 'h2' }
                elseif ($size -ge ($script:BaseFontSize + 1)) { $kind = 'h3' }
            }
            $table = $child.SelectSingleNode('./*[local-name()="Table"]')
            if ($null -ne $table) {
                $tableBlock = Convert-TableToBlock $table
                $tableBlock['objectId'] = $oid
                $tableBlock['outlineId'] = $script:CurrentOutlineId
                [void]$out.Add($tableBlock)
            }
            if ($text -ne '') {
                [void]$out.Add([ordered]@{ type = $kind; text = $text; level = $level; objectId = $oid; outlineId = $script:CurrentOutlineId })
            }
            foreach ($kids in $child.SelectNodes('./*[local-name()="OEChildren"]')) {
                $sub = Convert-ContainerToBlocks $kids ($level + 1)
                foreach ($item in $sub) { [void]$out.Add($item) }
            }
        } elseif ($localName -eq 'Table') {
            [void]$out.Add((Convert-TableToBlock $child))
        } elseif ($localName -eq 'OEChildren') {
            $sub = Convert-ContainerToBlocks $child $level
            foreach ($item in $sub) { [void]$out.Add($item) }
        } elseif ($localName -eq 'Image') {
            [void]$out.Add([ordered]@{ type = 'image'; text = ''; level = $level })
        } elseif ($localName -eq 'InkDrawing' -or $localName -eq 'InkParagraph') {
            [void]$out.Add([ordered]@{ type = 'ink'; text = ''; level = $level })
        } elseif ($localName -eq 'InsertedFile' -or $localName -eq 'MediaFile') {
            [void]$out.Add([ordered]@{ type = 'file'; text = ''; level = $level })
        }
    }
    return $out
}

function Convert-PageXmlToBlocks($doc) {
    $nm = New-NsMgr $doc
    $title = ''
    $titleNode = $doc.SelectSingleNode('//one:Page/one:Title', $nm)
    if ($null -ne $titleNode) { $title = (Get-TextOf $titleNode) }
    # Baseline for recognising formatted headings on the way back out.
    $script:BaseFontSize = Get-DocBaseFontSize $doc
    if ($script:BaseFontSize -le 0) { $script:BaseFontSize = Get-OneNoteDefaultFontSize }
    if ($script:BaseFontSize -le 0) { $script:BaseFontSize = 12.0 }
    $blocks = New-Object System.Collections.ArrayList
    foreach ($node in $doc.DocumentElement.ChildNodes) {
        if ($node.NodeType -ne 'Element') { continue }
        if ($node.LocalName -eq 'Outline') {
            $script:CurrentOutlineId = $node.GetAttribute('objectID')
            foreach ($children in $node.SelectNodes('./*[local-name()="OEChildren"]')) {
                $sub = Convert-ContainerToBlocks $children 0
                foreach ($item in $sub) { [void]$blocks.Add($item) }
            }
        }
    }
    return @{ title = $title; blocks = @($blocks) }
}

# ---------------------------------------------------------------------------
# Section / page resolution
# ---------------------------------------------------------------------------

function Walk-SectionNode($node, [string]$notebook, [string]$groupPath, $list) {
    foreach ($child in $node.ChildNodes) {
        if ($child.NodeType -ne 'Element') { continue }
        $localName = $child.LocalName
        if ($localName -eq 'Section') {
            [void]$list.Add([ordered]@{
                name        = $child.GetAttribute('name')
                id          = $child.GetAttribute('ID')
                path        = $child.GetAttribute('path')
                notebook    = $notebook
                group       = $groupPath
                isRecycle   = ($child.GetAttribute('isInRecycleBin') -eq 'true')
                isLocal     = $false
            })
        } elseif ($localName -eq 'Folder' -or $localName -eq 'SectionGroup') {
            $childName = $child.GetAttribute('name')
            $next = $childName
            if ($groupPath -ne '') { $next = $groupPath + '/' + $childName }
            Walk-SectionNode $child $notebook $next $list
        }
    }
}

function Resolve-ExtraSectionId($entry) {
    if ($null -eq $entry) { return '' }
    $path = ''
    if ($entry.PSObject.Properties.Name -contains 'path') { $path = [string]$entry.path }
    if ($path -ne '') {
        # Standalone sections get a fresh OneNote ID after every restart, so the
        # reliable key is the .one file path, not a cached ID. OpenHierarchy on an
        # already-open section is idempotent and returns its current ID.
        # OpenHierarchy only accepts backslashes: normalise so a forward-slash
        # path in the config still resolves (otherwise the section looks missing).
        $path = $path.Trim().Replace('/', '\')
        $id = $null
        (Get-App).OpenHierarchy($path, '', [ref]$id)
        return [string]$id
    }
    return [string]$entry.id
}

function Get-AllSections($params) {
    $list = New-Object System.Collections.ArrayList
    $doc = New-XmlDoc (Get-HierarchyXml 3 '')
    foreach ($notebook in $doc.DocumentElement.SelectNodes('./*[local-name()="Notebook"]')) {
        Walk-SectionNode $notebook $notebook.GetAttribute('name') '' $list
    }
    foreach ($extra in @($params.extraSections)) {
        if ($null -eq $extra) { continue }
        $extraId = Resolve-ExtraSectionId $extra
        if ($extraId -eq '') { continue }
        $extraPath = ''
        if ($extra.PSObject.Properties.Name -contains 'path') { $extraPath = [string]$extra.path }
        [void]$list.Add([ordered]@{
            name      = [string]$extra.name
            id        = $extraId
            path      = $extraPath
            notebook  = ''
            group     = ''
            isRecycle = $false
            isLocal   = $true
        })
    }
    return $list
}

function Resolve-SectionId($params) {
    if ($null -ne $params.sectionId -and "$($params.sectionId)" -ne '') {
        return [string]$params.sectionId
    }
    $wantName = [string]$params.sectionName
    $wantNotebook = [string]$params.notebook
    if ($wantName -eq '') {
        Fail 'BAD_ARGUMENT' 'sectionId or sectionName is required'
    }
    $sections = Get-AllSections $params
    $exact = @()
    $fuzzy = @()
    foreach ($section in $sections) {
        if ($section.isRecycle) { continue }
        if ($wantNotebook -ne '' -and $section.notebook -ne $wantNotebook) { continue }
        if ($section.name -eq $wantName) { $exact += $section; continue }
        if ($section.name.ToLower().Contains($wantName.ToLower())) { $fuzzy += $section }
    }
    $pick = $null
    if ($exact.Count -ge 1) { $pick = $exact[0] }
    elseif ($fuzzy.Count -ge 1) { $pick = $fuzzy[0] }
    if ($null -eq $pick) {
        Fail 'SECTION_NOT_FOUND' ('section not found: ' + $wantName)
    }
    return [string]$pick.id
}

function Get-SectionPagesXml([string]$sectionId) {
    $xml = Get-HierarchyXml 4 $sectionId
    if ($null -eq $xml -or $xml -eq '') { Fail 'SECTION_NOT_FOUND' ('cannot read section: ' + $sectionId) }
    return (New-XmlDoc $xml)
}

function Find-PageInSection($doc, [string]$pageTitle) {
    $exact = $null
    $fuzzy = $null
    foreach ($page in $doc.SelectNodes('//*[local-name()="Page"]')) {
        $name = $page.GetAttribute('name')
        if ($name -eq $pageTitle) { $exact = $page; break }
        if ($null -eq $fuzzy -and $name.ToLower().Contains($pageTitle.ToLower())) { $fuzzy = $page }
    }
    if ($null -ne $exact) { return $exact }
    return $fuzzy
}

function Resolve-PageId($params) {
    if ($null -ne $params.pageId -and "$($params.pageId)" -ne '') {
        return [string]$params.pageId
    }
    $title = [string]$params.pageTitle
    if ($title -eq '') {
        Fail 'BAD_ARGUMENT' 'pageId or pageTitle is required'
    }
    if (($null -ne $params.sectionId -and "$($params.sectionId)" -ne '') -or ($null -ne $params.sectionName -and "$($params.sectionName)" -ne '')) {
        $sectionId = Resolve-SectionId $params
        $doc = Get-SectionPagesXml $sectionId
        $page = Find-PageInSection $doc $title
        if ($null -eq $page) { Fail 'PAGE_NOT_FOUND' ('page not found in section: ' + $title) }
        return $page.GetAttribute('ID')
    }
    $doc = New-XmlDoc (Get-HierarchyXml 4 '')
    $page = Find-PageInSection $doc $title
    if ($null -eq $page) { Fail 'PAGE_NOT_FOUND' ('page not found: ' + $title) }
    return $page.GetAttribute('ID')
}

# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------

function Get-PageEntries($sectionNode, [bool]$includeRecycle) {
    $pages = New-Object System.Collections.ArrayList
    foreach ($page in $sectionNode.SelectNodes('./*[local-name()="Page"]')) {
        $inRecycle = ($page.GetAttribute('isInRecycleBin') -eq 'true')
        if ($inRecycle -and -not $includeRecycle) { continue }
        [void]$pages.Add([ordered]@{
            name         = $page.GetAttribute('name')
            id           = $page.GetAttribute('ID')
            pageLevel    = $page.GetAttribute('pageLevel')
            lastModified = $page.GetAttribute('lastModifiedTime')
        })
    }
    return $pages
}

function Action-List($params) {
    $includePages = $true
    if ($null -ne $params.includePages) { $includePages = [bool]$params.includePages }
    $includeRecycle = $false
    if ($null -ne $params.includeRecycleBin) { $includeRecycle = [bool]$params.includeRecycleBin }
    if ($null -ne $params.includeRecycle) { $includeRecycle = [bool]$params.includeRecycle }
    $wantNotebook = [string]$params.notebook
    $wantSection = [string]$params.section
    $maxSections = 0
    if ($null -ne $params.maxSections) { $maxSections = [int]$params.maxSections }
    if ($maxSections -lt 0) { $maxSections = 0 }
    $maxPages = 0
    if ($null -ne $params.maxPages) { $maxPages = [int]$params.maxPages }
    if ($maxPages -lt 0) { $maxPages = 0 }
    $totalSections = 0
    $totalPages = 0

    $doc = New-XmlDoc (Get-HierarchyXml 4 '')
    $notebooks = New-Object System.Collections.ArrayList
    foreach ($notebook in $doc.DocumentElement.SelectNodes('./*[local-name()="Notebook"]')) {
        $notebookName = $notebook.GetAttribute('name')
        $notebookId = $notebook.GetAttribute('ID')
        if ($wantNotebook -ne '' -and $notebookName -ne $wantNotebook -and $notebookId -ne $wantNotebook) {
            if (-not $notebookName.ToLower().Contains($wantNotebook.ToLower())) { continue }
        }
        $sections = New-Object System.Collections.ArrayList
        $flat = New-Object System.Collections.ArrayList
        Walk-SectionNode $notebook $notebookName '' $flat
        foreach ($section in $flat) {
            if ($section.isRecycle -and -not $includeRecycle) { continue }
            if ($wantSection -ne '') {
                if ($section.name -ne $wantSection -and -not $section.name.ToLower().Contains($wantSection.ToLower())) { continue }
            }
            $pages = @()
            if ($includePages) {
                $sectionNode = $null
                foreach ($candidate in $notebook.SelectNodes('.//*[local-name()="Section"]')) {
                    if ($candidate.GetAttribute('ID') -eq $section.id) { $sectionNode = $candidate; break }
                }
                if ($null -ne $sectionNode) {
                    $pages = @(Get-PageEntries $sectionNode $includeRecycle)
                }
                $totalPages += $pages.Count
                if ($maxPages -gt 0 -and $pages.Count -gt $maxPages) {
                    $pages = @($pages[0..($maxPages - 1)])
                }
            }
            $totalSections++
            [void]$sections.Add([ordered]@{
                name      = $section.name
                id        = $section.id
                path      = $section.path
                group     = $section.group
                isRecycle = $section.isRecycle
                pages     = @($pages)
            })
            if ($maxSections -gt 0 -and $sections.Count -ge $maxSections) { break }
        }
        if ($sections.Count -gt 0) {
            [void]$notebooks.Add([ordered]@{
                name     = $notebookName
                id       = $notebookId
                path     = $notebook.GetAttribute('path')
                sections = @($sections)
            })
        }
    }

    $localSections = New-Object System.Collections.ArrayList
    foreach ($extra in @($params.extraSections)) {
        if ($null -eq $extra) { continue }
        $extraId = ''
        try { $extraId = Resolve-ExtraSectionId $extra } catch { }
        $extraPath = ''
        if ($extra.PSObject.Properties.Name -contains 'path') { $extraPath = [string]$extra.path }
        $entry = [ordered]@{
            name  = [string]$extra.name
            id    = $extraId
            path  = $extraPath
            pages = @()
            error = $null
        }
        if ($extraId -eq '') {
            $entry.error = 'cannot resolve local section (missing id or path)'
            [void]$localSections.Add($entry)
            continue
        }
        try {
            $doc2 = Get-SectionPagesXml $extraId
            $sectionNode = $doc2.SelectSingleNode('//*[local-name()="Section"]')
            if ($null -ne $sectionNode) {
                $entry.path = $sectionNode.GetAttribute('path')
                $entry.pages = @(Get-PageEntries $sectionNode $includeRecycle)
                if ($maxPages -gt 0 -and $entry.pages.Count -gt $maxPages) {
                    $entry.pages = @($entry.pages[0..($maxPages - 1)])
                }
            }
        } catch {
            $entry.error = $_.Exception.Message
        }
        [void]$localSections.Add($entry)
    }

    $truncated = $false
    if (($maxSections -gt 0 -and $totalSections -gt $maxSections) -or ($maxPages -gt 0 -and $totalPages -gt $maxPages)) {
        $truncated = $true
    }
    return [ordered]@{
        notebooks     = @($notebooks)
        localSections = @($localSections)
        counts        = [ordered]@{ notebooks = $notebooks.Count; sections = $totalSections; pages = $totalPages }
        truncated     = $truncated
    }
}

function Action-Read($params) {
    $pageId = Resolve-PageId $params
    $xml = ''
    (Get-App).GetPageContent($pageId, [ref]$xml, 1)
    $doc = New-XmlDoc $xml
    $converted = Convert-PageXmlToBlocks $doc
    $blocks = @($converted.blocks)
    $maxBlocks = 0
    if ($null -ne $params.maxBlocks) { $maxBlocks = [int]$params.maxBlocks }
    if ($maxBlocks -gt 0 -and $blocks.Count -gt $maxBlocks) {
        $blocks = @($blocks[0..($maxBlocks - 1)])
        $blocks += [ordered]@{ type = 'note'; text = ('... ' + $maxBlocks + ' of ' + $converted.blocks.Count + ' blocks shown (truncated)'); level = 0 }
    }
    return [ordered]@{
        pageId       = $pageId
        title        = $converted.title
        lastModified = $doc.DocumentElement.GetAttribute('lastModifiedTime')
        blockCount   = $converted.blocks.Count
        truncated    = ($blocks.Count -gt 0 -and $converted.blocks.Count -gt $blocks.Count)
        blocks       = $blocks
    }
}

function Action-Search($params) {
    $query = [string]$params.query
    if ($query -eq '') { Fail 'BAD_ARGUMENT' 'query is required' }
    $limit = 20
    if ($null -ne $params.limit) { $limit = [int]$params.limit }
    if ($limit -le 0) { $limit = 20 }
    $scopeId = ''
    if ($null -ne $params.scopeId -and "$($params.scopeId)" -ne '') { $scopeId = [string]$params.scopeId }
    $includeRecycle = $false
    if ($null -ne $params.includeRecycleBin) { $includeRecycle = [bool]$params.includeRecycleBin }
    $includeContent = $false
    if ($null -ne $params.includeContent) { $includeContent = [bool]$params.includeContent }
    $maxScan = 200
    if ($null -ne $params.maxScan) { $maxScan = [int]$params.maxScan }
    if ($maxScan -le 0) { $maxScan = 200 }
    $needle = $query.ToLower()

    # OneNote's FindPages only works when the Windows Search index is present.
    # On this machine it ignores the query and returns the whole hierarchy, so
    # the plugin does its own scan: page titles always, page text on request.
    $doc = New-XmlDoc (Get-HierarchyXml 4 $scopeId)
    $results = New-Object System.Collections.ArrayList
    $scanned = 0
    foreach ($section in $doc.SelectNodes('//*[local-name()="Section"]')) {
        if ($results.Count -ge $limit) { break }
        if (-not $includeRecycle -and $section.GetAttribute('isInRecycleBin') -eq 'true') { continue }
        $notebookName = ''
        $parent = $section.ParentNode
        while ($null -ne $parent) {
            if ($parent.LocalName -eq 'Notebook') { $notebookName = $parent.GetAttribute('name'); break }
            $parent = $parent.ParentNode
        }
        foreach ($page in $section.SelectNodes('./*[local-name()="Page"]')) {
            if ($results.Count -ge $limit) { break }
            if (-not $includeRecycle -and $page.GetAttribute('isInRecycleBin') -eq 'true') { continue }
            $pageId = $page.GetAttribute('ID')
            $pageName = $page.GetAttribute('name')
            $matched = $pageName.ToLower().Contains($needle)
            $snippet = ''
            if (-not $matched -and $includeContent -and $scanned -lt $maxScan) {
                $scanned++
                try {
                    $content = ''
                    (Get-App).GetPageContent($pageId, [ref]$content, 0)
                    $text = Convert-OneNoteText ([regex]::Replace($content, '<[^>]*>', ' '))
                    $hit = $text.ToLower().IndexOf($needle)
                    if ($hit -ge 0) {
                        $matched = $true
                        $from = [Math]::Max(0, $hit - 40)
                        $take = [Math]::Min(140, $text.Length - $from)
                        $snippet = ([regex]::Replace($text.Substring($from, $take), '\s+', ' ')).Trim()
                    }
                } catch {
                    # page is locked/encrypted - skip it quietly
                }
            }
            if ($matched) {
                [void]$results.Add([ordered]@{
                    pageId   = $pageId
                    title    = $pageName
                    notebook = $notebookName
                    section  = $section.GetAttribute('name')
                    modified = $page.GetAttribute('lastModifiedTime')
                    snippet  = $snippet
                })
            }
        }
    }
    $method = 'title'
    if ($includeContent) { $method = 'title+content' }
    return [ordered]@{
        query   = $query
        method  = $method
        results = @($results)
        count   = $results.Count
        scanned = $scanned
    }
}
function Action-Create($params) {
    $sectionId = Resolve-SectionId $params
    $title = [string]$params.title
    $pageColor = [string]$params.pageColor
    $pageId = $null
    (Get-App).CreateNewPage($sectionId, [ref]$pageId, 0)
    $chunks = Invoke-WritePageContent $pageId $title $params.blocks $pageColor $params.pageSize $sectionId
    return [ordered]@{ pageId = $pageId; title = $title; sectionId = $sectionId; chunks = $chunks; pageColor = $pageColor; baseFontSize = $script:BaseFontSize }
}

function Action-Append($params) {
    $pageId = Resolve-PageId $params
    $title = ''
    if ($null -ne $params.title) { $title = [string]$params.title }
    $pageColor = [string]$params.pageColor
    $chunks = Invoke-WritePageContent $pageId $title $params.blocks $pageColor $params.pageSize ''
    return [ordered]@{ pageId = $pageId; appended = $true; chunks = $chunks }
}

# Page-level settings only (background color and/or page size) - content untouched.
function Action-SetSettings($params) {
    $pageId = Resolve-PageId $params
    $color = ''
    if ($null -ne $params.color) { $color = [string]$params.color }
    $pageSize = $params.pageSize
    if ($color -eq '' -and $null -eq $pageSize) { Fail 'BAD_ARGUMENT' 'color or pageSize is required' }
    $settings = Render-PageSettingsXml $color $pageSize
    if ($settings -eq '') { Fail 'BAD_ARGUMENT' 'nothing to apply (empty color and pageSize)' }
    $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '">' + $settings + '</one:Page>'
    (Get-App).UpdatePageContent($patch, 0)
    return [ordered]@{ pageId = $pageId; color = $color; pageSize = $pageSize }
}

function Action-SetColor($params) {
    $pageId = Resolve-PageId $params
    $color = [string]$params.color
    if ($color -eq '') { Fail 'BAD_ARGUMENT' 'color is required (e.g. #C7EDCC, automatic or none)' }
    $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '"><one:PageSettings RTL="false" color="' + (ConvertTo-Attr $color) + '"/></one:Page>'
    (Get-App).UpdatePageContent($patch, 0)
    return [ordered]@{ pageId = $pageId; color = $color }
}

function Action-Update($params) {
    $pageId = Resolve-PageId $params
    $app = Get-App
    $xml = ''
    $app.GetPageContent($pageId, [ref]$xml, 0)
    $doc = New-XmlDoc $xml
    $nm = New-NsMgr $doc
    $outlineIds = @()
    foreach ($outline in $doc.SelectNodes('//one:Page/one:Outline', $nm)) {
        $oid = $outline.GetAttribute('objectID')
        if ($oid -ne '') { $outlineIds += $oid }
    }
    foreach ($oid in $outlineIds) {
        $app.DeletePageContent($pageId, $oid, 0, $true)
    }
    $title = ''
    if ($null -ne $params.title) { $title = [string]$params.title }
    $pageColor = [string]$params.pageColor
    $chunks = Invoke-WritePageContent $pageId $title $params.blocks $pageColor $params.pageSize ''
    return [ordered]@{
        pageId          = $pageId
        updated         = $true
        removedOutlines = $outlineIds.Count
        chunks          = $chunks
    }
}

# ---------------------------------------------------------------------------
# Partial page editing
# ---------------------------------------------------------------------------
# UpdatePageContent is incremental: an element carrying an objectID replaces
# exactly that object and leaves the rest of the page alone. A partial edit
# therefore re-sends only the OUTLINE that owns the target block, with the
# change applied and every sibling copied back verbatim.

function Get-OwnerOutline($node) {
    $cur = $node
    while ($null -ne $cur) {
        if ($cur.LocalName -eq 'Outline') { return $cur }
        $cur = $cur.ParentNode
    }
    return $null
}

function New-CdataTextElement($doc, [string]$text) {
    $t = $doc.CreateElement('one', 'T', $script:NS)
    [void]$t.AppendChild($doc.CreateCDataSection($text))
    return $t
}

function Action-Edit($params) {
    $pageId = Resolve-PageId $params
    $mode = [string]$params.mode
    if ($mode -eq '') { Fail 'BAD_ARGUMENT' 'mode is required' }
    $objectId = ''
    if ($null -ne $params.objectId) { $objectId = [string]$params.objectId }
    $find = ''
    if ($null -ne $params.find) { $find = [string]$params.find }
    if ($objectId -eq '' -and $find -eq '') {
        Fail 'BAD_ARGUMENT' 'objectId or find is required (onenote_read with withIds=true gives objectIds)'
    }

    $app = Get-App
    $xml = ''
    $app.GetPageContent($pageId, [ref]$xml, 0)
    $doc = New-XmlDoc $xml
    $nm = New-NsMgr $doc

    $target = $null
    if ($objectId -ne '') {
        foreach ($cand in $doc.SelectNodes('//one:OE', $nm)) {
            if ($cand.GetAttribute('objectID') -eq $objectId) { $target = $cand; break }
        }
        if ($null -eq $target) { Fail 'NOT_FOUND' ('no outline element with objectID ' + $objectId) }
    } else {
        # Target by text: unique match required, so a partial edit can never hit
        # the wrong paragraph.
        $hits = @()
        foreach ($cand in $doc.SelectNodes('//one:OE', $nm)) {
            $text = Get-DirectText $cand
            if ($text -ne '' -and $text.Contains($find)) { $hits += $cand }
        }
        if ($hits.Count -eq 0) { Fail 'NOT_FOUND' ('no block contains: ' + $find) }
        if ($hits.Count -gt 1) { Fail 'AMBIGUOUS' ('find matches ' + $hits.Count + ' blocks; use objectId instead') }
        $target = $hits[0]
        $objectId = $target.GetAttribute('objectID')
    }
    $outline = Get-OwnerOutline $target
    if ($null -eq $outline) { Fail 'NOT_FOUND' 'target block has no owning outline' }
    $outlineId = $outline.GetAttribute('objectID')

    # Detached copy of just that outline; the source document is never mutated.
    $cloneDoc = New-XmlDoc ('<one:Page xmlns:one="' + $script:NS + '">' + $outline.OuterXml + '</one:Page>')
    $cnm = New-NsMgr $cloneDoc
    $clone = $cloneDoc.SelectSingleNode('//one:Outline', $cnm)
    $cloneTarget = $null
    foreach ($cand in $clone.SelectNodes('.//one:OE', $cnm)) {
        if ($cand.GetAttribute('objectID') -eq $objectId) { $cloneTarget = $cand; break }
    }
    if ($null -eq $cloneTarget) { Fail 'NOT_FOUND' 'target block missing from the outline copy' }

    $changed = 0
    $deleteOutline = $false

    if ($mode -eq 'replaceText') {
        $search = [string]$params.search
        $replacement = ''
        if ($null -ne $params.replacement) { $replacement = [string]$params.replacement }
        if ($search -eq '') { Fail 'BAD_ARGUMENT' 'search is required for replaceText' }
        $all = $false
        if ($null -ne $params.all) { $all = [bool]$params.all }
        foreach ($t in @($cloneTarget.SelectNodes('./*[local-name()="T"]'))) {
            $text = $t.InnerText
            $idx = $text.IndexOf($search)
            if ($idx -lt 0) { continue }
            if ($all) {
                $newText = $text.Replace($search, $replacement)
                $changed += ([regex]::Matches($text, [regex]::Escape($search))).Count
            } else {
                $newText = $text.Substring(0, $idx) + $replacement + $text.Substring($idx + $search.Length)
                $changed += 1
            }
            $t.RemoveAll()
            [void]$t.AppendChild($cloneDoc.CreateCDataSection($newText))
            if (-not $all) { break }
        }
        if ($changed -eq 0) { Fail 'NOT_FOUND' ('text not found in block ' + $objectId) }
    } elseif ($mode -eq 'replaceBlock') {
        $text = ''
        if ($null -ne $params.text) { $text = [string]$params.text }
        foreach ($t in @($cloneTarget.SelectNodes('./*[local-name()="T"]'))) { [void]$cloneTarget.RemoveChild($t) }
        [void]$cloneTarget.AppendChild((New-CdataTextElement $cloneDoc $text))
        $changed = 1
    } elseif ($mode -eq 'insertAfter') {
        $text = ''
        if ($null -ne $params.text) { $text = [string]$params.text }
        $kind = ''
        if ($null -ne $params.kind) { $kind = [string]$params.kind }
        $newOe = $cloneDoc.CreateElement('one', 'OE', $script:NS)
        if ($kind -eq 'h2' -or $kind -eq 'h3') {
            $base = Get-DocBaseFontSize $doc
            if ($base -le 0) { $base = 12.0 }
            $delta = 4
            if ($kind -eq 'h3') { $delta = 2 }
            $newOe.SetAttribute('style', ('font-size:' + (Format-PtSize ($base + $delta)) + ';'))
        } elseif ($kind -eq 'bullet' -or $kind -eq 'number') {
            # OneNote requires the bullet attribute; numbered entries are written
            # the same way as on create (bullet marker + "N. " prefix text).
            $list = $cloneDoc.CreateElement('one', 'List', $script:NS)
            $marker = $cloneDoc.CreateElement('one', 'Bullet', $script:NS)
            $marker.SetAttribute('bullet', $script:BulletValue)
            $marker.SetAttribute('fontSize', '11.0')
            [void]$list.AppendChild($marker)
            [void]$newOe.AppendChild($list)
            if ($kind -eq 'number') {
                $ordinal = 1
                if ($null -ne $params.ordinal) { $ordinal = [int]$params.ordinal }
                $text = [string]$ordinal + '. ' + $text
            }
        }
        [void]$newOe.AppendChild((New-CdataTextElement $cloneDoc $text))
        [void]$cloneTarget.ParentNode.InsertAfter($newOe, $cloneTarget)
        $changed = 1
    } elseif ($mode -eq 'deleteBlock') {
        [void]$cloneTarget.ParentNode.RemoveChild($cloneTarget)
        if ($clone.SelectNodes('.//one:OE', $cnm).Count -eq 0) { $deleteOutline = $true }
        $changed = 1
    } else {
        Fail 'BAD_ARGUMENT' ('unknown mode: ' + $mode)
    }

    if ($deleteOutline) {
        if ($outlineId -ne '') { $app.DeletePageContent($pageId, $outlineId, 0, $true) }
    } else {
        $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '">' + $clone.OuterXml + '</one:Page>'
        $app.UpdatePageContent($patch, 0)
    }
    return [ordered]@{
        pageId         = $pageId
        mode           = $mode
        objectId       = $objectId
        outlineId      = $outlineId
        changed        = $changed
        applied        = $true
        outlineRemoved = $deleteOutline
    }
}

function Action-Delete($params) {
    $pageId = Resolve-PageId $params
    $permanent = $false
    if ($null -ne $params.permanent) { $permanent = [bool]$params.permanent }
    (Get-App).DeleteHierarchy($pageId, 0, $permanent)
    return [ordered]@{ pageId = $pageId; deleted = $true; permanent = $permanent }
}

function Action-Publish($params) {
    $pageId = Resolve-PageId $params
    $target = [string]$params.path
    if ($target -eq '') { Fail 'BAD_ARGUMENT' 'path is required' }
    # OneNote's Publish rejects forward slashes with 0x80042006 ("invalid
    # hierarchy"); it only accepts native backslash paths.
    $target = $target.Replace('/', '\')
    if (-not [IO.Path]::IsPathRooted($target)) { Fail 'BAD_ARGUMENT' 'path must be absolute' }
    $formatName = 'one'
    if ($null -ne $params.format -and "$($params.format)" -ne '') { $formatName = ([string]$params.format).ToLower() }
    $formatCode = 0
    switch ($formatName) {
        'one'    { $formatCode = 0 }
        'onepkg' { $formatCode = 1 }
        'mht'    { $formatCode = 2 }
        'pdf'    { $formatCode = 3 }
        'xps'    { $formatCode = 4 }
        'docx'   { $formatCode = 5 }
        'doc'    { $formatCode = 5 }
        'emf'    { $formatCode = 6 }
        'html'   { $formatCode = 7 }
        default  { Fail 'BAD_ARGUMENT' ('unsupported format: ' + $formatName) }
    }
    $overwrite = $true
    if ($null -ne $params.overwrite) { $overwrite = [bool]$params.overwrite }
    if (Test-Path -LiteralPath $target) {
        if (-not $overwrite) { Fail 'FILE_EXISTS' ('target exists: ' + $target) }
        Remove-Item -LiteralPath $target -Force
    }
    $dir = [IO.Path]::GetDirectoryName($target)
    if ($dir -ne '' -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    (Get-App).Publish($pageId, $target, $formatCode)
    $size = 0
    if (Test-Path -LiteralPath $target) { $size = (Get-Item -LiteralPath $target).Length }
    return [ordered]@{ pageId = $pageId; path = $target; format = $formatName; bytes = $size }
}

# ---------------------------------------------------------------------------
# Sections
# ---------------------------------------------------------------------------

# A freshly created section always comes with one empty page. Remove it only
# when the section holds exactly one page AND that page has no body outlines,
# so user content is never touched.
function Remove-SectionBlankPage([string]$sectionId) {
    try {
        $xml = ''
        (Get-App).GetHierarchy($sectionId, 4, [ref]$xml)
        $doc = New-XmlDoc $xml
        $nm = New-NsMgr $doc
        $pages = @($doc.SelectNodes('//one:Page', $nm))
        if ($pages.Count -ne 1) { return $false }
        $pageId = [string]$pages[0].GetAttribute('ID')
        $content = ''
        (Get-App).GetPageContent($pageId, [ref]$content, 0)
        $pdoc = New-XmlDoc $content
        $pnm = New-NsMgr $pdoc
        $body = @($pdoc.SelectNodes('//one:Page/one:Outline//one:OE', $pnm)).Count
        if ($body -gt 0) { return $false }
        (Get-App).DeleteHierarchy($pageId, 0, $true)
        return $true
    } catch {
        return $false
    }
}

# Create or open a standalone section file. OpenHierarchy with cft=3
# (cftSection) creates the .one when it does not exist and is idempotent.
function Action-CreateSection($params) {
    $target = [string]$params.path
    if ($target -eq '') { Fail 'BAD_ARGUMENT' 'path is required (.one file path)' }
    $target = $target.Trim().Replace('/', '\')
    if (-not [IO.Path]::IsPathRooted($target)) { Fail 'BAD_ARGUMENT' 'path must be absolute' }
    $dir = [IO.Path]::GetDirectoryName($target)
    if ($dir -ne '' -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    $existed = Test-Path -LiteralPath $target
    $id = $null
    (Get-App).OpenHierarchy($target, '', [ref]$id, 3)
    $keepBlank = $false
    if ($null -ne $params.keepBlankPage) { $keepBlank = [bool]$params.keepBlankPage }
    $blankRemoved = $false
    if (-not $existed -and -not $keepBlank) { $blankRemoved = Remove-SectionBlankPage ([string]$id) }
    return [ordered]@{
        sectionId         = [string]$id
        path              = $target
        created           = (-not $existed)
        blankPageRemoved  = $blankRemoved
    }
}

# Rename a section. Notebook sections accept this; standalone sections are
# refused by OneNote with 0x8004200B (SECTION_READ_ONLY).
function Action-RenameSection($params) {
    $sectionId = ''
    if ($null -ne $params.sectionId) { $sectionId = [string]$params.sectionId }
    $path = ''
    if ($null -ne $params.path) { $path = ([string]$params.path).Trim().Replace('/', '\') }
    $name = ''
    if ($null -ne $params.name) { $name = [string]$params.name }
    if ($sectionId -eq '' -and $path -ne '') {
        $resolved = $null
        (Get-App).OpenHierarchy($path, '', [ref]$resolved)
        $sectionId = [string]$resolved
    }
    if ($sectionId -eq '') { Fail 'BAD_ARGUMENT' 'sectionId (or path) is required' }
    if ($name -eq '') { Fail 'BAD_ARGUMENT' 'name is required' }
    if ($name -match '[\\/:*?"<>|]') { Fail 'BAD_ARGUMENT' 'name contains characters OneNote rejects: \ / : * ? " < > |' }
    $xml = '<?xml version="1.0"?><one:Section xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $sectionId) + '" name="' + (ConvertTo-Attr $name) + '"/>'
    (Get-App).UpdateHierarchy($xml)
    return [ordered]@{ sectionId = $sectionId; name = $name; path = $path }
}
# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

$result = @{ ok = $false; code = 'UNKNOWN'; message = 'no action executed' }
try {
    # NOTE: never assign the parsed object to $Request/$Response - those
    # variables carry a [string] type constraint from the param block and would
    # silently coerce the whole request into its string representation.
    $payload = Read-JsonFile $Request
    # OneNote bullet glyph index: 0 = big black circle, 1 = U+2022, 2 = U+2022,
    # 3 = hollow circle, 6 = star, 7 = diamond, 20 = arrowhead.
    $script:BulletValue = '1'
    $script:BulletValues = @()
    $script:NestUnderHeadings = $false
    if ($null -ne $payload.params -and $null -ne $payload.params.nestUnderHeadings) {
        $script:NestUnderHeadings = [bool]$payload.params.nestUnderHeadings
    }
    if ($null -ne $payload.params) {
        if ($null -ne $payload.params.bulletValue) {
            $script:BulletValue = [string]$payload.params.bulletValue
        }
        if ($null -ne $payload.params.bulletValues) {
            $vals = @($payload.params.bulletValues)
            if ($vals.Count -gt 0) {
                $script:BulletValues = @($vals | ForEach-Object { [string]$_ })
                $script:BulletValue = $script:BulletValues[0]
            }
        }
    }
    $action = [string]$payload.action
    $params = $payload.params
    if ($null -eq $params) { $params = (New-Object PSObject) }
    switch ($action) {
        'list'    { $data = Action-List $params }
        'read'    { $data = Action-Read $params }
        'search'  { $data = Action-Search $params }
        'create'  { $data = Action-Create $params }
        'update'  { $data = Action-Update $params }
        'append'  { $data = Action-Append $params }
        'edit'    { $data = Action-Edit $params }
        'delete'  { $data = Action-Delete $params }
        'setcolor' { $data = Action-SetColor $params }
        'setsettings' { $data = Action-SetSettings $params }
        'createsection' { $data = Action-CreateSection $params }
        'renamesection' { $data = Action-RenameSection $params }
        'publish' { $data = Action-Publish $params }
        default   { Fail 'BAD_ACTION' ('unknown action: ' + $action) }
    }
    $result = @{ ok = $true; data = $data }
} catch {
    $message = $_.Exception.Message
    if ($message.StartsWith('__DSH__')) {
        # Business error raised by Fail(): a stable code we already own.
        $body = $message.Substring(7)
        $idx = $body.IndexOf('|')
        if ($idx -ge 0) {
            $code = $body.Substring(0, $idx)
            $message = $body.Substring($idx + 1)
        } else {
            $code = 'ONENOTE_ERROR'
            $message = $body
        }
        $result = @{ ok = $false; code = $code; message = $message }
    } else {
        # COM / RPC error: map the HRESULT to a readable code and hint.
        $info = Get-ComError $_.Exception
        $line = ''
        try {
            if ($_.InvocationInfo.ScriptLineNumber) { $line = ' [bridge line ' + $_.InvocationInfo.ScriptLineNumber + ']' }
        } catch { }
        $result = @{ ok = $false; code = $info.code; message = ($info.message + $line); hresult = $info.hresult }
    }
}

Write-JsonFile $Response $result
if (-not $result.ok) { exit 1 }
