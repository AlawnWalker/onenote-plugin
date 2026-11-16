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
# Column widths of the table being replaced by Action-Update; reused by
# Convert-RowsToTableXml so a rewrite keeps the user's manual layout.
$script:PreservedTableWidths = $null
# <one:Size> of the outlines being replaced by Action-Update (document order);
# Invoke-WritePageContent re-applies them so the block keeps the width.
$script:PreservedOutlineSizes = $null
# Where destructive actions drop a copy of the page XML (update / delete) so a
# wrong call can always be undone outside OneNote. Empty = %TEMP% default.
$script:BackupDir = ''
# Action-Update keeps images/ink/attachments by default; set false to drop them.
$script:PreserveMedia = $true
# Upper bound (pt) for images written from markdown when no width is given.
$script:ImageMaxWidth = 600.0
# Non-fatal notes collected while rendering (missing image files, dropped
# media, ...). Returned as `warnings` instead of failing the whole page.
$script:Warnings = New-Object System.Collections.ArrayList
# OneNote media elements one:OE can carry. Read-back reports them, and
# Action-Update carries them over the outline rewrite.
$script:MediaTags = @('Image', 'InkDrawing', 'InkParagraph', 'InsertedFile', 'MediaFile')

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
    '8004201A' = @{ code = 'INVALID_TARGET_PATH'; hint = 'OneNote rejected the export path - the extension must match the format (html -> .htm, mht -> .mht, docx -> .docx, pdf -> .pdf) and the directory must be writable' }
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
        # OneNote may move inline formatting into spans inside the CDATA (mixed
        # CJK/ASCII runs); inspect the text so code detection still sees the
        # monospace family and size.
        $inner = [string]$t.InnerText
        if ($inner -match 'font-(family|size):') { [void]$styles.Add($inner) }
    }
    foreach ($s in $styles) {
        $m = [regex]::Match([string]$s, 'font-size:\s*([0-9]+(?:\.[0-9]+)?)\s*pt')
        if ($m.Success) { return [double]$m.Groups[1].Value }
    }
    return 0
}

# Effective font family of one outline element (same precedence as font size).
# Used to recognise code blocks when codePrefix is empty: the writer pins them to
# a monospace family, so read-back does not have to rely on a text marker.
function Get-OeFontFamily($oe) {
    $styles = New-Object System.Collections.ArrayList
    if ($oe.HasAttribute('style')) { [void]$styles.Add($oe.GetAttribute('style')) }
    foreach ($t in $oe.SelectNodes('./*[local-name()="T"]')) {
        if ($t.HasAttribute('style')) { [void]$styles.Add($t.GetAttribute('style')) }
        # OneNote may move inline formatting into spans inside the CDATA (mixed
        # CJK/ASCII runs); inspect the text so code detection still sees the
        # monospace family and size.
        $inner = [string]$t.InnerText
        if ($inner -match 'font-(family|size):') { [void]$styles.Add($inner) }
    }
    foreach ($s in $styles) {
        $m = [regex]::Match([string]$s, 'font-family:\s*([^;"]+)')
        if ($m.Success) { return $m.Groups[1].Value.Trim().ToLowerInvariant() }
    }
    return ''
}

function Test-OeIsCode($oe) {
    $family = Get-OeFontFamily $oe
    if ($family -eq '') { return $false }
    return ($family -match 'consolas|courier|monaco|menlo|mono')
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
# Backups, media (images / ink / attachments) and image files
# ---------------------------------------------------------------------------

function Get-BackupDir {
    if ($script:BackupDir -ne '') { return $script:BackupDir }
    return (Join-Path $env:TEMP 'onenote-plugin-backups')
}

# Write a copy of the page XML before a destructive action. Returns the path, or
# '' when the copy could not be written (the action still runs: a full disk must
# not block a delete the caller asked for).
function Save-PageBackup([string]$pageId, [string]$reason, [string]$xml) {
    try {
        $dir = Get-BackupDir
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
        }
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss-fff')
        $short = ($pageId -replace '[^0-9A-Za-z]', '')
        if ($short.Length -gt 24) { $short = $short.Substring(0, 24) }
        $target = Join-Path $dir ($stamp + '-' + $reason + '-' + $short + '.xml')
        [IO.File]::WriteAllText($target, $xml, (New-Object System.Text.UTF8Encoding($false)))
        return $target
    } catch {
        return ''
    }
}

function Add-Warning([string]$text) {
    [void]$script:Warnings.Add($text)
}

function Get-OeMediaInfo($oe) {
    foreach ($tag in $script:MediaTags) {
        $node = $oe.SelectSingleNode('./*[local-name()="' + $tag + '"]')
        if ($null -eq $node) { continue }
        $kind = 'image'
        if ($tag -eq 'InkDrawing' -or $tag -eq 'InkParagraph') { $kind = 'ink' }
        elseif ($tag -eq 'InsertedFile' -or $tag -eq 'MediaFile') { $kind = 'file' }
        $width = 0.0
        $height = 0.0
        $size = $node.SelectSingleNode('./*[local-name()="Size"]')
        if ($null -ne $size) {
            try { $width = [double]$size.GetAttribute('width') } catch { }
            try { $height = [double]$size.GetAttribute('height') } catch { }
        }
        return [ordered]@{ type = $kind; tag = $tag; width = $width; height = $height }
    }
    return $null
}

# Detached, page-independent copy of a media OE: metadata and the page-local
# binary handles (CallbackID / OCRData) are dropped so the element can be sent
# to a different page without confusing OneNote.
function Get-PortableMediaXml($oe) {
    $clone = $oe.CloneNode($true)
    $volatile = @('objectID', 'author', 'authorInitials', 'authorResolutionID',
        'lastModifiedBy', 'lastModifiedByInitials', 'lastModifiedByResolutionID',
        'creationTime', 'lastModifiedTime')
    $stack = New-Object System.Collections.Stack
    $stack.Push($clone)
    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        if ($node.NodeType -ne 'Element') { continue }
        foreach ($attr in $volatile) {
            if ($node.HasAttribute($attr)) { [void]$node.RemoveAttribute($attr) }
        }
        for ($i = $node.ChildNodes.Count - 1; $i -ge 0; $i--) {
            $child = $node.ChildNodes[$i]
            if ($child.NodeType -eq 'Element' -and ($child.LocalName -eq 'CallbackID' -or $child.LocalName -eq 'OCRData')) {
                [void]$node.RemoveChild($child)
            }
        }
        foreach ($child in $node.ChildNodes) {
            if ($child.NodeType -eq 'Element') { $stack.Push($child) }
        }
    }
    return $clone.OuterXml
}

# Media elements of a page in document order, each with the text of the closest
# preceding paragraph (the anchor used to put it back in the same place).
function Get-MediaItems($doc) {
    $nm = New-NsMgr $doc
    $items = New-Object System.Collections.ArrayList
    $outlineIndex = -1
    foreach ($outline in $doc.SelectNodes('//one:Page/one:Outline', $nm)) {
        $outlineIndex++
        $lastText = ''
        foreach ($oe in $outline.SelectNodes('.//one:OE', $nm)) {
            $info = Get-OeMediaInfo $oe
            if ($null -ne $info) {
                [void]$items.Add([ordered]@{
                    xml          = (Get-PortableMediaXml $oe)
                    type         = [string]$info.type
                    width        = [double]$info.width
                    height       = [double]$info.height
                    anchor       = $lastText
                    outlineIndex = $outlineIndex
                })
                continue
            }
            $text = Get-DirectText $oe
            if ($text -ne '') { $lastText = $text }
        }
    }
    return ,$items
}

function Find-OutlineIndexByText($outlines, $nm, [string]$text) {
    if ($text -eq '') { return -1 }
    for ($i = 0; $i -lt $outlines.Count; $i++) {
        foreach ($oe in $outlines[$i].SelectNodes('.//one:OE', $nm)) {
            if ((Get-DirectText $oe) -eq $text) { return $i }
        }
    }
    return -1
}

# Put media OEs back after the text rewrite. Each item goes to the outline that
# still holds its anchor paragraph, else to the last outline; items with no
# outline at all land in a fresh one.
function Add-MediaToPage([string]$pageId, $items) {
    $list = @($items)
    if ($list.Count -eq 0) { return ,@() }
    $app = Get-App
    $xml = ''
    $app.GetPageContent($pageId, [ref]$xml, 0)
    $doc = New-XmlDoc $xml
    $nm = New-NsMgr $doc
    $outlines = @($doc.SelectNodes('//one:Page/one:Outline', $nm))
    $placed = New-Object System.Collections.ArrayList

    if ($outlines.Count -eq 0) {
        $body = ''
        foreach ($item in $list) { $body += [string]$item.xml }
        $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '"><one:Outline><one:OEChildren>' + $body + '</one:OEChildren></one:Outline></one:Page>'
        $app.UpdatePageContent($patch, 0)
        foreach ($item in $list) {
            [void]$placed.Add([ordered]@{ type = $item.type; width = $item.width; height = $item.height; anchor = $item.anchor; placed = 'new-outline' })
        }
        return ,@($placed)
    }

    # Map every item to one outline, then send one patch per outline.
    $plan = @{}
    foreach ($item in $list) {
        $target = Find-OutlineIndexByText $outlines $nm ([string]$item.anchor)
        if ($target -lt 0) { $target = $outlines.Count - 1 }
        if (-not $plan.ContainsKey($target)) { $plan[$target] = New-Object System.Collections.ArrayList }
        [void]$plan[$target].Add($item)
    }
    foreach ($target in @($plan.Keys | Sort-Object)) {
        $clone = $outlines[$target].CloneNode($true)
        $container = $clone.SelectSingleNode('./*[local-name()="OEChildren"]')
        if ($null -eq $container) {
            $container = $doc.CreateElement('one', 'OEChildren', $script:NS)
            [void]$clone.AppendChild($container)
        }
        # Several media items can share one anchor (two pictures under the same
        # paragraph). Remember the last node inserted for each anchor so the
        # second one lands behind the first instead of in front of it.
        $lastInserted = @{}
        foreach ($item in $plan[$target]) {
            $wrapper = New-XmlDoc ('<one:Outline xmlns:one="' + $script:NS + '"><one:OEChildren>' + [string]$item.xml + '</one:OEChildren></one:Outline>')
            # Put the media back next to the paragraph it followed before the
            # rewrite; the anchor is only the outline fallback when that text is
            # gone.
            $key = [string]$item.anchor
            $anchorNode = $null
            if ($key -ne '' -and $lastInserted.ContainsKey($key)) {
                $anchorNode = $lastInserted[$key]
            } elseif ($key -ne '') {
                foreach ($cand in $clone.SelectNodes('.//*[local-name()="OE"]')) {
                    if ((Get-DirectText $cand) -eq $key) { $anchorNode = $cand; break }
                }
            }
            $usedAnchor = $false
            foreach ($node in @($wrapper.DocumentElement.SelectNodes('./*[local-name()="OEChildren"]/*[local-name()="OE"]'))) {
                $imported = $doc.ImportNode($node, $true)
                if ($null -ne $anchorNode) {
                    [void]$anchorNode.ParentNode.InsertAfter($imported, $anchorNode)
                    $usedAnchor = $true
                } else {
                    [void]$container.AppendChild($imported)
                }
                $anchorNode = $imported
            }
            if ($key -ne '' -and $null -ne $anchorNode) { $lastInserted[$key] = $anchorNode }
            [void]$placed.Add([ordered]@{
                type   = $item.type
                width  = $item.width
                height = $item.height
                anchor = $item.anchor
                placed = $(if ($usedAnchor) { 'after anchor' } else { 'outline ' + $target + ' (end)' })
            })
        }
        $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '">' + $clone.OuterXml + '</one:Page>'
        $app.UpdatePageContent($patch, 0)
        Start-Sleep -Milliseconds 300
    }
    return ,@($placed)
}

# Pixel dimensions of a PNG/JPEG/GIF/BMP file; 0x0 when unknown.
function Get-ImagePixelSize([string]$path) {
    try {
        $bytes = [IO.File]::ReadAllBytes($path)
    } catch {
        return @(0, 0)
    }
    if ($null -eq $bytes -or $bytes.Length -lt 26) { return @(0, 0) }
    # PNG: signature then IHDR width/height (big endian).
    if ($bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) {
        $w = ([int64]$bytes[16] -shl 24) -bor ([int64]$bytes[17] -shl 16) -bor ([int64]$bytes[18] -shl 8) -bor [int64]$bytes[19]
        $h = ([int64]$bytes[20] -shl 24) -bor ([int64]$bytes[21] -shl 16) -bor ([int64]$bytes[22] -shl 8) -bor [int64]$bytes[23]
        return @([int]$w, [int]$h)
    }
    # GIF: 'GIF8' then little-endian width/height.
    if ($bytes[0] -eq 0x47 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46) {
        return @(([int]$bytes[6] + ([int]$bytes[7] -shl 8)), ([int]$bytes[8] + ([int]$bytes[9] -shl 8)))
    }
    # BMP: 40-byte header, little-endian width/height at 18/22.
    if ($bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4D) {
        return @(([int]$bytes[18] + ([int]$bytes[19] -shl 8)), ([int]$bytes[22] + ([int]$bytes[23] -shl 8)))
    }
    # JPEG: walk the markers looking for a start-of-frame segment.
    if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8) {
        $i = 2
        while ($i -lt $bytes.Length - 9) {
            if ($bytes[$i] -ne 0xFF) { $i++; continue }
            $marker = [int]$bytes[$i + 1]
            if ($marker -eq 0xD8 -or $marker -eq 0xD9 -or ($marker -ge 0xD0 -and $marker -le 0xD7) -or $marker -eq 0x01) { $i += 2; continue }
            $len = ([int]$bytes[$i + 2] -shl 8) + [int]$bytes[$i + 3]
            if ($marker -ge 0xC0 -and $marker -le 0xCF -and $marker -ne 0xC4 -and $marker -ne 0xC8 -and $marker -ne 0xCC) {
                $h = ([int]$bytes[$i + 5] -shl 8) + [int]$bytes[$i + 6]
                $w = ([int]$bytes[$i + 7] -shl 8) + [int]$bytes[$i + 8]
                return @($w, $h)
            }
            if ($len -le 0) { break }
            $i += 2 + $len
        }
    }
    return @(0, 0)
}

# Build the one:OE that embeds an image file. Width in points; the height keeps
# the aspect ratio. Returns '' (and records a warning) when the file is unusable.
function Get-ImageOeFromFile([string]$path, $requestedWidth) {
    if ($path -eq '') {
        Add-Warning 'image block without a source path'
        return ''
    }
    $full = $path.Trim().Replace('/', '\')
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
        Add-Warning ('image file not found: ' + $full)
        return ''
    }
    $pixels = Get-ImagePixelSize $full
    $pw = [double]$pixels[0]
    $ph = [double]$pixels[1]
    if ($pw -le 0 -or $ph -le 0) { $pw = 480.0; $ph = 320.0 }
    $width = 0.0
    if ($null -ne $requestedWidth) {
        try { $width = [double]$requestedWidth } catch { $width = 0.0 }
    }
    if ($width -le 0) { $width = [Math]::Min($pw, $script:ImageMaxWidth) }
    $height = [Math]::Round($width * $ph / $pw, 1)
    try {
        $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($full))
    } catch {
        Add-Warning ('cannot read image file: ' + $full)
        return ''
    }
    $oeId = '{' + [guid]::NewGuid().ToString().ToUpper() + '}{1}{B0}'
    $imgId = '{' + [guid]::NewGuid().ToString().ToUpper() + '}{2}{B0}'
    return '<one:OE objectID="' + $oeId + '" alignment="left"><one:Image objectID="' + $imgId + '"><one:Size width="' + $width + '" height="' + $height + '"/><one:Data>' + $b64 + '</one:Data></one:Image></one:OE>'
}

# ---------------------------------------------------------------------------
# Page XML construction from the plugin's block model
# ---------------------------------------------------------------------------

# Heading paragraphs are real formatted paragraphs: font size plus paragraph
# spacing. Only the ASCII style string is used (no font-family) so this script
# stays ASCII-only; OneNote merges the size into the page's default font.
# Sizes are RELATIVE to the body size (see Resolve-BaseFontSize) and are
# configurable (`headingDelta: [h2, h3]`, default [2, 1]): a 12pt notebook gets
# 14/13 by default. Older pages were written with [4, 2] (16/14); read-back
# recognises both candidate sizes.
$script:HeadingSpace = @{
    h2 = @{ before = '8.0'; after = '5.0' }
    h3 = @{ before = '5.0'; after = '5.0' }
}
$script:HeadingDelta = @{ h2 = 2; h3 = 1 }
$script:BaseFontSize = 12.0
# List bullet glyph size: 0 = follow the body size; > 0 = explicit pt.
$script:BulletFontSize = 0.0

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
    # Reuse the replaced table's column widths when the column count is
    # unchanged; otherwise fall back to equal widths (680pt total).
    $widths = @()
    if ($null -ne $script:PreservedTableWidths -and @($script:PreservedTableWidths).Count -eq $ncols) {
        $widths = @($script:PreservedTableWidths)
    } else {
        $width = [Math]::Round(680.0 / $ncols, 1)
        for ($i = 0; $i -lt $ncols; $i++) { $widths += $width }
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<one:Table bordersVisible="true" hasHeaderRow="true"><one:Columns>')
    for ($i = 0; $i -lt $ncols; $i++) {
        [void]$sb.Append('<one:Column index="' + $i + '" width="' + $widths[$i] + '"/>')
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
            # "<br>" separates the lines of one cell: write one OE per line, or
            # OneNote mangles the literal tag and drops the text around it.
            [void]$sb.Append('<one:Cell><one:OEChildren>')
            foreach ($part in ($cell -split '(?i)<br\s*/?>')) {
                $line = $part.Replace('>', [string][char]0xFF1E)
                [void]$sb.Append('<one:OE><one:T>' + (ConvertTo-CData $line) + '</one:T></one:OE>')
            }
            [void]$sb.Append('</one:OEChildren></one:Cell>')
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
        if ($type -eq 'image') {
            # Markdown `![alt](path)`: the file is read here (the bridge runs on
            # the same machine as the page), so the model never has to move
            # base64 through the host.
            $imageXml = Get-ImageOeFromFile ([string]$node.src) $node.width
            if ($imageXml -ne '') {
                [void]$sb.Append($imageXml)
            } else {
                # Pin the placeholder to the body size: an unstyled OE written
                # under a heading inherits the heading's font and reads back as
                # a phantom `## [missing image]`.
                $note = '[missing image] ' + [string]$node.src
                [void]$sb.Append('<one:OE style="font-size:' + (Format-PtSize $script:BaseFontSize) + '"><one:T>' + (ConvertTo-CData $note) + '</one:T></one:OE>')
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
        # one:Bullet fontSize takes a bare number, not "Npt".
        $bulletSizeXml = '{0:0.0}' -f $(if ($script:BulletFontSize -gt 0) { $script:BulletFontSize } else { $script:BaseFontSize })
        if ($type -eq 'bullet') {
            $listXml = '<one:List><one:Bullet bullet="' + $levelBullet + '" fontSize="' + $bulletSizeXml + '"/></one:List>'
        } elseif ($type -eq 'number') {
            $listXml = '<one:List><one:Bullet bullet="' + $levelBullet + '" fontSize="' + $bulletSizeXml + '"/></one:List>'
            $ordinal = 1
            if ($null -ne $node.ordinal -and "$($node.ordinal)" -ne '') { $ordinal = [int]$node.ordinal }
            $text = "$ordinal. " + $text
        }
        $headingXml = ''
        if ($type -eq 'para' -or $type -eq 'bullet' -or $type -eq 'number') {
            # Pin body/list paragraphs to the resolved body size: a child nested
            # under a heading would otherwise inherit the heading's larger font and
            # be mistaken for another heading on read-back.
            $headingXml = ' style="font-size:' + (Format-PtSize $script:BaseFontSize) + '"'
        }
        if ($script:HeadingDelta.ContainsKey($type)) {
            $delta = [int]$script:HeadingDelta[$type]
            $space = $script:HeadingSpace[$type]
            $headingXml = ' style="font-size:' + (Format-PtSize ($script:BaseFontSize + $delta)) + '" spaceBefore="' + $space.before + '" spaceAfter="' + $space.after + '"'
        } elseif ($type -eq 'code' -or $type -eq 'tree') {
            # Pin code lines to the body size and a monospace family. Without an
            # explicit size they inherit the surrounding paragraph style (a heading
            # written just above makes them look like h2/h3 on read-back), and the
            # font family is what read-back uses to recognise a code block when
            # codePrefix is empty.
            $headingXml = ' style="font-size:' + (Format-PtSize $script:BaseFontSize) + ';font-family:Consolas"'
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
            src      = $block.src
            width    = $block.width
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
    # Re-apply the replaced page's outline sizes; the index-wise mapping keeps
    # chunked pages intact, and an index without a stored size stays auto.
    if ($null -ne $script:PreservedOutlineSizes) {
        for ($i = 0; $i -lt $outlines.Count; $i++) {
            if ($i -lt @($script:PreservedOutlineSizes).Count -and $script:PreservedOutlineSizes[$i] -ne '') {
                $outlines[$i] = $outlines[$i].Replace('<one:Outline>', '<one:Outline>' + $script:PreservedOutlineSizes[$i])
            }
        }
    }
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
            [void]$cells.Add(($texts -join '<br>'))
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
            if ($kind -eq 'para') {
                # Code blocks written by this plugin carry a monospace family and
                # the body size; recognise them before the heading heuristics so a
                # heading above cannot turn them into h2/h3 on the way back out.
                # Empty code lines keep their OE, so classify them before the
                # empty-text guard as well.
                if (Test-OeIsCode $child) {
                    $kind = 'code'
                } elseif ($text -ne '') {
                    # Formatted headings written by this plugin sit at
                    # base + headingDelta; pages written by older versions used
                    # base + 4 / base + 2. Accept both candidate sizes, then fall
                    # back to the legacy thresholds for anything in between.
                    $size = Get-OeFontSize $child
                    $h2s = @(($script:BaseFontSize + [double]$script:HeadingDelta['h2']), ($script:BaseFontSize + 4.0))
                    $h3s = @(($script:BaseFontSize + [double]$script:HeadingDelta['h3']), ($script:BaseFontSize + 2.0))
                    $isH2 = $false
                    $isH3 = $false
                    foreach ($cand in $h2s) { if ([math]::Abs($size - $cand) -lt 0.26) { $isH2 = $true } }
                    foreach ($cand in $h3s) { if ([math]::Abs($size - $cand) -lt 0.26) { $isH3 = $true } }
                    if ($isH2) { $kind = 'h2' }
                    elseif ($isH3) { $kind = 'h3' }
                    elseif ($size -ge ($script:BaseFontSize + 3)) { $kind = 'h2' }
                    elseif ($size -ge ($script:BaseFontSize + 1)) { $kind = 'h3' }
                }
            }
            $table = $child.SelectSingleNode('./*[local-name()="Table"]')
            # Images, ink and attachments live inside their own OE and carry no
            # text: report them so a caller can see that the page has media
            # (and so a read/update round trip can prove nothing was lost).
            $media = Get-OeMediaInfo $child
            if ($null -ne $media) {
                [void]$out.Add([ordered]@{
                    type      = [string]$media.type
                    text      = ''
                    level     = $level
                    objectId  = $oid
                    outlineId = $script:CurrentOutlineId
                    width     = [double]$media.width
                    height    = [double]$media.height
                })
            }
            if ($null -ne $table) {
                $tableBlock = Convert-TableToBlock $table
                $tableBlock['objectId'] = $oid
                $tableBlock['outlineId'] = $script:CurrentOutlineId
                [void]$out.Add($tableBlock)
            }
            if ($text -ne '' -or $kind -eq 'code') {
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
    try {
        # piBasic (0): everything except the binary payloads. Images still show
        # up as one:Image/one:Size, which is all read-back needs, and a page
        # with big pictures stays cheap to read.
        (Get-App).GetPageContent($pageId, [ref]$xml, 0)
    } catch {
        # PowerShell wraps the COM failure: the OneNote HRESULT is on an inner
        # exception, so walk the chain instead of reading HResult directly.
        $hex = Get-DeepHResult $_.Exception
        if ($hex -eq '80042014' -or $hex -eq '80042005') {
            Fail 'PAGE_DELETED' ('page not found or in the recycle bin: ' + $pageId + '. OneNote COM cannot read or restore recycled pages; restore it in the OneNote UI, or rebuild it from the backupPath that onenote_update / onenote_delete returned')
        }
        throw
    }
    $doc = New-XmlDoc $xml
    $converted = Convert-PageXmlToBlocks $doc
    $blocks = @($converted.blocks)
    $maxBlocks = 0
    if ($null -ne $params.maxBlocks) { $maxBlocks = [int]$params.maxBlocks }
    if ($maxBlocks -gt 0 -and $blocks.Count -gt $maxBlocks) {
        $blocks = @($blocks[0..($maxBlocks - 1)])
        $blocks += [ordered]@{ type = 'note'; text = ('... ' + $maxBlocks + ' of ' + $converted.blocks.Count + ' blocks shown (truncated)'); level = 0 }
    }
    # Media inventory: what a caller must not lose when rewriting the page, and
    # the only way to notice that a rewrite dropped a picture.
    $images = New-Object System.Collections.ArrayList
    $inkCount = 0
    $fileCount = 0
    foreach ($block in $converted.blocks) {
        $kind = [string]$block.type
        if ($kind -eq 'ink') { $inkCount++ }
        elseif ($kind -eq 'file') { $fileCount++ }
        if ($kind -eq 'image' -or $kind -eq 'ink' -or $kind -eq 'file') {
            [void]$images.Add([ordered]@{
                type      = $kind
                width     = $block.width
                height    = $block.height
                objectId  = $block.objectId
                outlineId = $block.outlineId
            })
        }
    }
    # OneNote's own lastModifiedTime does not advance for API writes (verified:
    # UpdatePageContent changes the content while the attribute keeps the page
    # creation time), so a content hash is the reliable change detector.
    $sha = [System.Security.Cryptography.SHA1]::Create()
    $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($xml))
    $revision = ([BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
    if ($revision.Length -gt 16) { $revision = $revision.Substring(0, 16) }
    return [ordered]@{
        pageId       = $pageId
        title        = $converted.title
        lastModified = $doc.DocumentElement.GetAttribute('lastModifiedTime')
        revision     = $revision
        blockCount   = $converted.blocks.Count
        truncated    = ($blocks.Count -gt 0 -and $converted.blocks.Count -gt $blocks.Count)
        blocks       = $blocks
        images       = @($images)
        media        = [ordered]@{ images = ($images.Count - $inkCount - $fileCount); ink = $inkCount; files = $fileCount; total = $images.Count }
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
    return [ordered]@{ pageId = $pageId; title = $title; sectionId = $sectionId; chunks = $chunks; pageColor = $pageColor; baseFontSize = $script:BaseFontSize; warnings = @($script:Warnings) }
}

function Action-Append($params) {
    $pageId = Resolve-PageId $params
    $title = ''
    if ($null -ne $params.title) { $title = [string]$params.title }
    $pageColor = [string]$params.pageColor
    $chunks = Invoke-WritePageContent $pageId $title $params.blocks $pageColor $params.pageSize ''
    return [ordered]@{ pageId = $pageId; appended = $true; chunks = $chunks; warnings = @($script:Warnings) }
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
    # piBinaryData (1): the image/attachment payloads must travel with the
    # backup and with the media that is re-inserted after the text rewrite.
    $app.GetPageContent($pageId, [ref]$xml, 1)
    $doc = New-XmlDoc $xml
    $nm = New-NsMgr $doc

    # Always keep a copy of the page before the outlines are deleted: OneNote's
    # own recycle bin is not readable through COM, so this file is the only way
    # back. A failed backup never blocks the update the caller asked for.
    $backupPath = ''
    $wantBackup = $true
    if ($null -ne $params.backup) { $wantBackup = [bool]$params.backup }
    if ($wantBackup) { $backupPath = Save-PageBackup $pageId 'update' $xml }

    # Media preservation: capture images / ink / attachments (with their anchor
    # text) before the outlines that own them are deleted. The capture always
    # happens, even with preserveMedia=false, so the result can report exactly
    # what the rewrite dropped instead of losing it silently.
    $preserveMedia = $script:PreserveMedia
    if ($null -ne $params.preserveMedia) { $preserveMedia = [bool]$params.preserveMedia }
    $mediaItems = Get-MediaItems $doc

    # Remember the existing table's column widths so the rewritten table keeps
    # the user's manual layout (Convert-RowsToTableXml reuses them).
    $script:PreservedTableWidths = @()
    $oldTable = $doc.SelectSingleNode('//*[local-name()="Table"]')
    if ($null -ne $oldTable) {
        foreach ($col in $oldTable.SelectNodes('./*[local-name()="Columns"]/*[local-name()="Column"]')) {
            $w = $col.GetAttribute('width')
            if ($w -ne '') {
                try { $script:PreservedTableWidths += [double]$w } catch { }
            }
        }
    }

    # Keep each outline's <one:Size> (width/height/isSetByUser) so the rewrite
    # does not fall back to OneNote's automatic layout.
    $script:PreservedOutlineSizes = @()
    $outlineIds = @()
    foreach ($outline in $doc.SelectNodes('//one:Page/one:Outline', $nm)) {
        $oid = $outline.GetAttribute('objectID')
        if ($oid -ne '') { $outlineIds += $oid }
        $size = $outline.SelectSingleNode('./one:Size', $nm)
        if ($null -ne $size) { $script:PreservedOutlineSizes += $size.OuterXml } else { $script:PreservedOutlineSizes += '' }
    }
    foreach ($oid in $outlineIds) {
        $app.DeletePageContent($pageId, $oid, 0, $true)
    }
    $title = ''
    if ($null -ne $params.title) { $title = [string]$params.title }
    $pageColor = [string]$params.pageColor
    $chunks = Invoke-WritePageContent $pageId $title $params.blocks $pageColor $params.pageSize ''
    $script:PreservedTableWidths = $null
    $script:PreservedOutlineSizes = $null

    # Put the media back and verify it actually landed on the page.
    $preserved = New-Object System.Collections.ArrayList
    if ($mediaItems.Count -gt 0) {
        if ($preserveMedia) {
            $preserved = Add-MediaToPage $pageId $mediaItems
            # OneNote applies the patch asynchronously; give it a moment before
            # deciding that a picture did not make it back.
            $landed = 0
            for ($attempt = 0; $attempt -lt 5; $attempt++) {
                Start-Sleep -Milliseconds 350
                $check = ''
                try { $app.GetPageContent($pageId, [ref]$check, 0) } catch { $check = '' }
                if ($check -eq '') { continue }
                $landedItems = Get-MediaItems (New-XmlDoc $check)
                $landed = $landedItems.Count
                if ($landed -ge $mediaItems.Count) { break }
            }
            if ($landed -lt $mediaItems.Count) {
                Add-Warning ('restored ' + $landed + ' of ' + $mediaItems.Count + ' media items; the page copy from the backup still has all of them')
            }
        } else {
            Add-Warning ('preserveMedia=false: ' + $mediaItems.Count + ' media item(s) were removed by this update')
            foreach ($item in $mediaItems) {
                [void]$preserved.Add([ordered]@{ type = $item.type; width = $item.width; height = $item.height; anchor = $item.anchor; placed = 'dropped' })
            }
        }
    }
    $after = ''
    $revision = ''
    try {
        $app.GetPageContent($pageId, [ref]$after, 0)
    } catch { $after = '' }
    if ($after -ne '') {
        $sha = [System.Security.Cryptography.SHA1]::Create()
        $revision = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($after)))).Replace('-', '').ToLowerInvariant()
        if ($revision.Length -gt 16) { $revision = $revision.Substring(0, 16) }
    }
    return [ordered]@{
        pageId          = $pageId
        updated         = $true
        removedOutlines = $outlineIds.Count
        chunks          = $chunks
        backupPath      = $backupPath
        preservedMedia  = @($preserved)
        mediaCount      = $mediaItems.Count
        revision        = $revision
        warnings        = @($script:Warnings)
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
    $exact = $false
    if ($null -ne $params.exact) { $exact = [bool]$params.exact }
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
        # the wrong paragraph. Substring by default; exact=true compares the
        # whole trimmed block text (ordinal, case-sensitive).
        $hits = @()
        foreach ($cand in $doc.SelectNodes('//one:OE', $nm)) {
            $text = Get-DirectText $cand
            if ($exact) {
                if ([string]::Equals($text.Trim(), $find.Trim(), [StringComparison]::Ordinal)) { $hits += $cand }
            } elseif ($text -ne '' -and $text.Contains($find)) {
                $hits += $cand
            }
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

# Insert one image file into an existing page. The target is an anchor block
# (objectId or unique text) or the end of the page; only the outline that owns
# the anchor is re-sent, so the rest of the page is untouched.
function Action-InsertImage($params) {
    $pageId = Resolve-PageId $params
    $path = ''
    if ($null -ne $params.path) { $path = [string]$params.path }
    if ($path -eq '') { Fail 'BAD_ARGUMENT' 'path is required (absolute path of a local image file)' }
    $path = $path.Trim().Replace('/', '\')
    if (-not [IO.Path]::IsPathRooted($path)) { Fail 'BAD_ARGUMENT' 'path must be absolute' }

    $requestedWidth = $null
    if ($null -ne $params.width) { $requestedWidth = $params.width }
    $oeXml = Get-ImageOeFromFile $path $requestedWidth
    if ($oeXml -eq '') { Fail 'IMAGE_NOT_FOUND' ('cannot read image file: ' + $path) }
    $newId = [regex]::Match($oeXml, 'objectID="([^"]+)"').Groups[1].Value
    $sizeMatch = [regex]::Match($oeXml, '<one:Size width="([^"]+)" height="([^"]+)"')
    # Numbers, not the regex strings: callers compare/format these.
    $widthValue = [double]$sizeMatch.Groups[1].Value
    $heightValue = [double]$sizeMatch.Groups[2].Value

    $app = Get-App
    $xml = ''
    $app.GetPageContent($pageId, [ref]$xml, 1)
    $doc = New-XmlDoc $xml
    $nm = New-NsMgr $doc

    # OneNote assigns its own objectID to a new OE, so remember the IDs that
    # already exist and diff after the patch: the reported objectId is the one
    # the caller can actually use with onenote_edit / insert_image anchors.
    $knownIds = @{}
    foreach ($oe in $doc.SelectNodes('//one:OE', $nm)) {
        $oid = $oe.GetAttribute('objectID')
        if ($oid -ne '') { $knownIds[$oid] = $true }
    }

    $backupPath = ''
    $wantBackup = $true
    if ($null -ne $params.backup) { $wantBackup = [bool]$params.backup }
    if ($wantBackup) { $backupPath = Save-PageBackup $pageId 'insertimage' $xml }

    $wrapperXml = '<one:Outline xmlns:one="' + $script:NS + '"><one:OEChildren>' + $oeXml + '</one:OEChildren></one:Outline>'
    $outlines = @($doc.SelectNodes('//one:Page/one:Outline', $nm))

    if ($outlines.Count -eq 0) {
        $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '"><one:Outline><one:OEChildren>' + $oeXml + '</one:OEChildren></one:Outline></one:Page>'
        $app.UpdatePageContent($patch, 0)
        Start-Sleep -Milliseconds 300
        $createdId = $newId
        $createdXml = ''
        try { $app.GetPageContent($pageId, [ref]$createdXml, 0) } catch { $createdXml = '' }
        if ($createdXml -ne '') {
            $createdDoc = New-XmlDoc $createdXml
            $cnm = New-NsMgr $createdDoc
            foreach ($oe in $createdDoc.SelectNodes('//one:OE', $cnm)) {
                $oid = $oe.GetAttribute('objectID')
                if ($oid -ne '' -and $null -ne $oe.SelectSingleNode('./*[local-name()="Image"]')) { $createdId = $oid; break }
            }
        }
        return [ordered]@{
            pageId = $pageId; objectId = $createdId; width = $widthValue; height = $heightValue
            position = 'new-outline'; path = $path; backupPath = $backupPath; warnings = @($script:Warnings)
        }
    }

    # Resolve the anchor block (objectId wins over text).
    $anchor = $null
    $anchorId = ''
    $anchorFromText = $false
    if ($null -ne $params.objectId -and "$($params.objectId)" -ne '') {
        $wantId = [string]$params.objectId
        foreach ($cand in $doc.SelectNodes('//one:OE', $nm)) {
            if ($cand.GetAttribute('objectID') -eq $wantId) { $anchor = $cand; break }
        }
        if ($null -eq $anchor) { Fail 'NOT_FOUND' ('no block with objectID ' + $wantId) }
    } elseif ($null -ne $params.find -and "$($params.find)" -ne '') {
        $needle = [string]$params.find
        $hits = @()
        foreach ($cand in $doc.SelectNodes('//one:OE', $nm)) {
            $text = Get-DirectText $cand
            if ($text -ne '' -and $text.Contains($needle)) { $hits += $cand }
        }
        if ($hits.Count -eq 0) { Fail 'NOT_FOUND' ('no block contains: ' + $needle) }
        if ($hits.Count -gt 1) { Fail 'AMBIGUOUS' ('find matches ' + $hits.Count + ' blocks; use objectId instead') }
        $anchor = $hits[0]
        $anchorFromText = $true
    }
    if ($null -ne $anchor) { $anchorId = $anchor.GetAttribute('objectID') }

    $position = 'after'
    if ($null -ne $params.position -and "$($params.position)" -ne '') { $position = ([string]$params.position).ToLower() }

    if ($null -eq $anchor) {
        # No anchor: append to the last outline.
        $owner = $outlines[$outlines.Count - 1]
        $clone = $owner.CloneNode($true)
        $container = $clone.SelectSingleNode('./*[local-name()="OEChildren"]')
        if ($null -eq $container) {
            $container = $doc.CreateElement('one', 'OEChildren', $script:NS)
            [void]$clone.AppendChild($container)
        }
        $newNode = $doc.ImportNode((New-XmlDoc $wrapperXml).DocumentElement.SelectSingleNode('.//*[local-name()="OE"]'), $true)
        [void]$container.AppendChild($newNode)
        $position = 'end'
    } else {
        $owner = Get-OwnerOutline $anchor
        if ($null -eq $owner) { Fail 'NOT_FOUND' 'anchor block has no owning outline' }
        $clone = $owner.CloneNode($true)
        $cloneAnchor = $null
        foreach ($cand in $clone.SelectNodes('.//*[local-name()="OE"]')) {
            if ($cand.GetAttribute('objectID') -eq $anchorId) { $cloneAnchor = $cand; break }
        }
        $newNode = $doc.ImportNode((New-XmlDoc $wrapperXml).DocumentElement.SelectSingleNode('.//*[local-name()="OE"]'), $true)
        if ($null -eq $cloneAnchor) {
            $container = $clone.SelectSingleNode('./*[local-name()="OEChildren"]')
            [void]$container.AppendChild($newNode)
            $position = 'end'
        } elseif ($position -eq 'before') {
            [void]$cloneAnchor.ParentNode.InsertBefore($newNode, $cloneAnchor)
        } else {
            $insertAfter = $cloneAnchor
            if ($anchorFromText) {
                # Text anchors are usually a heading/paragraph that gets several
                # pictures. Step over the media already sitting right behind it
                # so repeated "after this paragraph" calls line up in call order
                # instead of pushing the previous picture down each time. An
                # objectId anchor stays strict: insert right next to that block.
                $probe = $cloneAnchor.NextSibling
                while ($null -ne $probe) {
                    if ($probe.NodeType -eq 'Element') {
                        if ($probe.LocalName -ne 'OE' -or $null -eq (Get-OeMediaInfo $probe)) { break }
                        $insertAfter = $probe
                    }
                    $probe = $probe.NextSibling
                }
            }
            [void]$insertAfter.ParentNode.InsertAfter($newNode, $insertAfter)
        }
    }

    $patch = '<?xml version="1.0"?><one:Page xmlns:one="' + $script:NS + '" ID="' + (ConvertTo-Attr $pageId) + '">' + $clone.OuterXml + '</one:Page>'
    $app.UpdatePageContent($patch, 0)
    Start-Sleep -Milliseconds 300

    # Report the ID OneNote actually assigned: the generated one is not what
    # later reads/edits see. Existing OEs keep their IDs on an outline rewrite,
    # so the only image OE with an unknown ID is the one just inserted.
    $actualId = $newId
    $afterXml = ''
    try { $app.GetPageContent($pageId, [ref]$afterXml, 0) } catch { $afterXml = '' }
    if ($afterXml -ne '') {
        $afterDoc = New-XmlDoc $afterXml
        $anm = New-NsMgr $afterDoc
        foreach ($oe in $afterDoc.SelectNodes('//one:OE', $anm)) {
            $oid = $oe.GetAttribute('objectID')
            if ($oid -eq '' -or $knownIds.ContainsKey($oid)) { continue }
            if ($null -ne $oe.SelectSingleNode('./*[local-name()="Image"]')) { $actualId = $oid; break }
        }
    }

    return [ordered]@{
        pageId     = $pageId
        objectId   = $actualId
        width      = $widthValue
        height     = $heightValue
        position   = $position
        anchorId   = $anchorId
        path       = $path
        backupPath = $backupPath
        warnings   = @($script:Warnings)
    }
}

function Action-Delete($params) {
    $pageId = Resolve-PageId $params
    $permanent = $false
    if ($null -ne $params.permanent) { $permanent = [bool]$params.permanent }
    # Deleted pages cannot be read back through COM (0x80042014) and COM has no
    # undelete, so keep the XML first: this file is the only recovery path.
    $backupPath = ''
    $wantBackup = $true
    if ($null -ne $params.backup) { $wantBackup = [bool]$params.backup }
    if ($wantBackup) {
        try {
            $xml = ''
            (Get-App).GetPageContent($pageId, [ref]$xml, 1)
            if ($xml -ne '') { $backupPath = Save-PageBackup $pageId 'delete' $xml }
        } catch { $backupPath = '' }
    }
    (Get-App).DeleteHierarchy($pageId, 0, $permanent)
    return [ordered]@{
        pageId     = $pageId
        deleted    = $true
        permanent  = $permanent
        backupPath = $backupPath
        warnings   = @($script:Warnings)
    }
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
    # OneNote's Publish validates the extension against the chosen format; a
    # mismatched extension fails with 0x8004201A even when the file does not
    # exist (html must be .htm, not .html). Normalise instead of failing.
    $extensionByFormat = @{
        one = '.one'; onepkg = '.onepkg'; mht = '.mht'; pdf = '.pdf'
        xps = '.xps'; docx = '.docx'; doc = '.docx'; emf = '.emf'; html = '.htm'
    }
    $wantedExtension = $extensionByFormat[$formatName]
    if ($null -ne $wantedExtension) {
        $currentExtension = [IO.Path]::GetExtension($target)
        if ($currentExtension.ToLowerInvariant() -ne $wantedExtension) {
            $target = [IO.Path]::ChangeExtension($target, $wantedExtension)
        }
    }
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
    $script:BulletFontSize = 0.0
    if ($null -ne $payload.params -and $null -ne $payload.params.nestUnderHeadings) {
        $script:NestUnderHeadings = [bool]$payload.params.nestUnderHeadings
    }
    if ($null -ne $payload.params) {
        # Backup + media switches (host config, overridable per call).
        if ($null -ne $payload.params.backupDir) {
            $script:BackupDir = ([string]$payload.params.backupDir).Trim()
        }
        if ($null -ne $payload.params.preserveMedia) {
            $script:PreserveMedia = [bool]$payload.params.preserveMedia
        }
        if ($null -ne $payload.params.imageMaxWidth) {
            try {
                $mw = [double]$payload.params.imageMaxWidth
                if ($mw -gt 0) { $script:ImageMaxWidth = $mw }
            } catch { }
        }
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
        # Heading sizes are relative to the body: headingDelta = [h2, h3].
        if ($null -ne $payload.params.headingDelta) {
            $hd = $payload.params.headingDelta
            if ($null -ne $hd.h2) { try { $script:HeadingDelta['h2'] = [double]$hd.h2 } catch { } }
            if ($null -ne $hd.h3) { try { $script:HeadingDelta['h3'] = [double]$hd.h3 } catch { } }
        }
        # Bullet glyph size: "body" (default) follows the body size, a number is pt.
        if ($null -ne $payload.params.bulletFontSize) {
            $bf = ([string]$payload.params.bulletFontSize).Trim()
            if ($bf -ne '' -and $bf.ToLower() -ne 'body') {
                try {
                    $v = [double]$bf
                    if ($v -gt 0) { $script:BulletFontSize = $v }
                } catch { }
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
        'insertimage' { $data = Action-InsertImage $params }
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
