<#
.SYNOPSIS
    Reads a VCF Implementation Plan (.docx) and builds an Excel execution tracker
    with three tabs: Dashboard, Tracker and Issues Log.

.DESCRIPTION
    No Python, no Excel, no Word and no PowerShell modules are needed. The .docx is
    read as a zip, and the .xlsx is written directly as XML.

    What it takes from the Word file
      - The "Preparation" checklist table                 -> Pre-check rows
      - "General Requirements" / "Item + Requirements"    -> Pre-check rows
      - Every "Execution Plan" table (Action, Impact to Environment, Required,
        VM Downtime)                                      -> one phase each, one row per
                                                             action. Actions starting with
                                                             "Validate" become Checkpoints.
                                                             Required = No is marked "(optional)".

    Two ways to use it

      1. One file:
           .\plan_to_tracker.ps1 -Docx "My Plan.docx"
           .\plan_to_tracker.ps1 -Docx "My Plan.docx" -Output "My Tracker.xlsx"

      2. Drop folder:
           .\plan_to_tracker.ps1 -Watch "C:\Trackers\DropHere"
           .\plan_to_tracker.ps1 -Watch .\DropHere -Once

         Put a .docx in the folder. The tracker appears in <folder>\Trackers and the Word
         file moves to <folder>\Processed. A file that cannot be read moves to
         <folder>\Failed with a .txt note saying why. An existing tracker is never
         overwritten (it will hold the team's progress); a new run gets a timestamp in
         its name instead.

.PARAMETER Docx
    Word plan to convert.

.PARAMETER Output
    Output .xlsx path (single file mode). Defaults to "<plan name> - Tracker.xlsx" next to the plan.

.PARAMETER Watch
    Drop folder to watch for .docx files. Created if it does not exist.

.PARAMETER Once
    With -Watch: process what is there, then stop.

.PARAMETER Interval
    Seconds between checks in watch mode. Default 5.
#>
param(
    [string]$Docx,
    [string]$Output,
    [string]$Watch,
    [switch]$Once,
    [int]$Interval = 5
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

$script:Inv  = [System.Globalization.CultureInfo]::InvariantCulture
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:NsW  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

function Clean-Text([string]$t) {
    if ($null -eq $t) { return '' }
    $t = $t.Replace("`t", ' ')
    return ([regex]::Replace($t, '\s+', ' ')).Trim()
}

function Strip-See([string]$t) {
    # Drop "See <link text>." sentences, which mean little once the link is gone.
    $t = [regex]::Replace($t, '(^|\s)See [^.]*\.', ' ')
    return (Clean-Text $t)
}

function Esc-Xml([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Col-Name([int]$n) {
    $s = ''
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = [string][char](65 + $m) + $s
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $s
}

function Get-UniquePath([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $path }
    $dir  = Split-Path -Parent $path
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $ext  = [System.IO.Path]::GetExtension($path)
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss', $script:Inv)
    return (Join-Path $dir ($stem + ' ' + $stamp + $ext))
}

# ---------------------------------------------------------------------------
# Reading the Word plan
# ---------------------------------------------------------------------------

function Open-ZipRead([string]$path) {
    $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
    } catch {
        $fs.Dispose()
        throw
    }
    return $zip
}

function Read-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Get-ParaText($p, $nsm) {
    $sb = New-Object System.Text.StringBuilder
    $nodes = $p.SelectNodes('.//w:t | .//w:tab | .//w:br | .//w:cr', $nsm)
    foreach ($n in $nodes) {
        if ($n.LocalName -eq 't') { [void]$sb.Append($n.InnerText) }
        else { [void]$sb.Append(' ') }
    }
    return (Clean-Text $sb.ToString())
}

function Get-StyleName($p, $nsm, $styleMap) {
    $ps = $p.SelectSingleNode('w:pPr/w:pStyle', $nsm)
    if ($null -eq $ps) { return 'Normal' }
    $id = $ps.GetAttribute('val', $script:NsW)
    if ($styleMap.ContainsKey($id)) { return $styleMap[$id] }
    return $id
}

function Get-CellParas($tc, $nsm) {
    # Paragraph texts of one table cell, dropping empty ones and "Note ..." paragraphs.
    $out = New-Object System.Collections.ArrayList
    foreach ($p in $tc.SelectNodes('w:p', $nsm)) {
        $t = Get-ParaText $p $nsm
        if (-not $t) { continue }
        if ($t.ToLower().StartsWith('note ')) { continue }
        $t = Strip-See $t
        if ($t) { [void]$out.Add($t) }
    }
    return ,$out
}

function Get-RowCells($tr, $nsm) {
    # One entry per grid column, so a merged cell repeats like it does in Word's own object model.
    $cells = New-Object System.Collections.ArrayList
    foreach ($tc in $tr.SelectNodes('w:tc', $nsm)) {
        $parts = New-Object System.Collections.ArrayList
        foreach ($p in $tc.SelectNodes('w:p', $nsm)) {
            $t = Get-ParaText $p $nsm
            if ($t) { [void]$parts.Add($t) }
        }
        $text = Clean-Text ($parts -join ' ')
        $span = 1
        $gs = $tc.SelectSingleNode('w:tcPr/w:gridSpan', $nsm)
        if ($null -ne $gs) { $span = [int]$gs.GetAttribute('val', $script:NsW) }
        for ($k = 0; $k -lt $span; $k++) { [void]$cells.Add($text) }
    }
    return ,$cells
}

function Get-NumberFormats($numText) {
    # Maps "numId|ilvl" to the numbering format (decimal, bullet, ...), so bullets are not mistaken for steps.
    $map = @{}
    if (-not $numText) { return $map }
    $nd = New-Object System.Xml.XmlDocument
    $nd.LoadXml($numText)
    $nn = New-Object System.Xml.XmlNamespaceManager($nd.NameTable)
    $nn.AddNamespace('w', $script:NsW)
    $abs = @{}
    foreach ($a in $nd.SelectNodes('//w:abstractNum', $nn)) {
        $aid = $a.GetAttribute('abstractNumId', $script:NsW)
        foreach ($lv in $a.SelectNodes('w:lvl', $nn)) {
            $f = $lv.SelectSingleNode('w:numFmt', $nn)
            if ($null -ne $f) { $abs[$aid + '|' + $lv.GetAttribute('ilvl', $script:NsW)] = $f.GetAttribute('val', $script:NsW) }
        }
    }
    foreach ($n in $nd.SelectNodes('//w:num', $nn)) {
        $nid = $n.GetAttribute('numId', $script:NsW)
        $ab = $n.SelectSingleNode('w:abstractNumId', $nn)
        if ($null -eq $ab) { continue }
        $aid = $ab.GetAttribute('val', $script:NsW)
        for ($l = 0; $l -le 8; $l++) {
            $k = $aid + '|' + $l
            if ($abs.ContainsKey($k)) { $map[$nid + '|' + $l] = $abs[$k] }
        }
    }
    return $map
}

function Get-StepItems([string]$secName, $recs, [bool]$isTop) {
    # recs: Kind (num / lit / plain), Lvl, Text. Top-level steps become rows; deeper ones are counted as sub-steps.
    $minLvl = 99
    foreach ($r in $recs) { if ($r.Kind -eq 'num' -and $r.Lvl -lt $minLvl) { $minLvl = $r.Lvl } }
    $items = New-Object System.Collections.ArrayList
    $inLit = $false
    foreach ($r in $recs) {
        if ($r.Kind -eq 'lit') {
            [void]$items.Add(@{ Text = $r.Text; Subs = 0 }); $inLit = $true
        } elseif ($r.Kind -eq 'num') {
            if ($items.Count -gt 0 -and ($inLit -or $r.Lvl -gt $minLvl)) { $items[$items.Count - 1].Subs++ }
            else { [void]$items.Add(@{ Text = $r.Text; Subs = 0 }); $inLit = $false }
        }
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($it in $items) {
        $txt = $it.Text.TrimEnd(':').Trim()
        if (-not $txt) { continue }
        if ($txt.Length -gt 200) {
            $cut = $txt.Substring(0, 200)
            $sp = $cut.LastIndexOf(' ')
            if ($sp -gt 120) { $cut = $cut.Substring(0, $sp) }
            $txt = $cut + ' ...'
        }
        $kind = 'Task'
        if ($txt -match '^(?i)(validate|verify|confirm|check)\b') { $kind = 'Checkpoint' }
        $note = $secName
        if ($it.Subs -gt 0) { $note += '. ' + $it.Subs + ' sub-steps in the source document' }
        [void]$out.Add(@{ Kind = $kind; Text = $txt; Note = $note })
    }
    if ($out.Count -eq 0 -and $isTop) {
        [void]$out.Add(@{ Kind = 'Task'; Text = ('Complete: ' + $secName); Note = $secName })
    }
    return ,$out
}

function Read-ProcedurePlan($body, $nsm, $styleMap, $numFmt) {
    # For procedure documents (for example a Broadcom export) that have no Execution Plan tables:
    # one phase per main heading, one row per top-level numbered step.
    $phases = New-Object System.Collections.ArrayList
    $secName = $null; $secTop = $false
    $recs = New-Object System.Collections.ArrayList
    $flush = {
        if ($null -ne $secName) {
            $its = Get-StepItems $secName $recs $secTop
            if ($its.Count -gt 0) { [void]$phases.Add(@{ Name = $secName; Items = $its }) }
        }
    }
    foreach ($child in $body.ChildNodes) {
        if ($child.NodeType -ne 'Element' -or $child.LocalName -ne 'p') { continue }
        $text = Get-ParaText $child $nsm
        if (-not $text) { continue }
        $style = (Get-StyleName $child $nsm $styleMap).ToLower()
        if ($style.StartsWith('heading')) {
            # Headings that contain a tab are table header lines, not section titles.
            if ($null -ne $child.SelectSingleNode('.//w:tab', $nsm)) { continue }
            & $flush
            $secName = $text; $secTop = ($style -match '1$')
            $recs = New-Object System.Collections.ArrayList
            continue
        }
        if ($null -eq $secName) { continue }
        $numPr = $child.SelectSingleNode('w:pPr/w:numPr', $nsm)
        $isNum = $false; $lvl = 0
        if ($null -ne $numPr) {
            $idn = $numPr.SelectSingleNode('w:numId', $nsm)
            $ln = $numPr.SelectSingleNode('w:ilvl', $nsm)
            if ($null -ne $ln) { $lvl = [int]$ln.GetAttribute('val', $script:NsW) }
            $nid = ''; if ($null -ne $idn) { $nid = $idn.GetAttribute('val', $script:NsW) }
            if ($nid -ne '' -and $nid -ne '0') {
                $isNum = $true
                $k = $nid + '|' + $lvl
                if ($numFmt.ContainsKey($k) -and $numFmt[$k] -eq 'bullet') { $isNum = $false }
            }
        }
        $clean = Strip-See $text
        if ($isNum) {
            if ($clean) { [void]$recs.Add(@{ Kind = 'num'; Lvl = $lvl; Text = $clean }) }
        } elseif ($text -match '^\d+\.\s+\S') {
            [void]$recs.Add(@{ Kind = 'lit'; Lvl = 0; Text = (Strip-See ([regex]::Replace($text, '^\d+\.\s*', ''))) })
        }
    }
    & $flush
    return ,$phases
}

function Read-Plan([string]$path) {
    $zip = $null
    try {
        $zip = Open-ZipRead $path
    } catch {
        $ex = $_.Exception
        $bad = ($ex -is [System.IO.InvalidDataException])
        if (-not $bad -and $null -ne $ex.InnerException) { $bad = ($ex.InnerException -is [System.IO.InvalidDataException]) }
        if ($bad) { throw 'This is not a valid .docx file. It may be an old .doc file (save it as .docx) or a damaged download.' }
        throw
    }
    try {
        $docText = Read-ZipText $zip 'word/document.xml'
        $styText = Read-ZipText $zip 'word/styles.xml'
        $numText = Read-ZipText $zip 'word/numbering.xml'
    } finally {
        $zip.Dispose()
    }
    if (-not $docText) {
        throw 'This is not a valid .docx file (no word/document.xml inside). It may be an old .doc file saved with the wrong extension.'
    }

    $nsm = $null
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($docText)
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('w', $script:NsW)

    $styleMap = @{}
    if ($styText) {
        $sdoc = New-Object System.Xml.XmlDocument
        $sdoc.LoadXml($styText)
        $snsm = New-Object System.Xml.XmlNamespaceManager($sdoc.NameTable)
        $snsm.AddNamespace('w', $script:NsW)
        foreach ($st in $sdoc.SelectNodes('//w:style', $snsm)) {
            $id = $st.GetAttribute('styleId', $script:NsW)
            $nm = $st.SelectSingleNode('w:name', $snsm)
            if ($id -and $null -ne $nm) { $styleMap[$id] = $nm.GetAttribute('val', $script:NsW) }
        }
    }

    $body = $doc.SelectSingleNode('/w:document/w:body', $nsm)
    if ($null -eq $body) { throw 'This is not a valid .docx file (no document body).' }

    $title = ''
    foreach ($child in $body.ChildNodes) {
        if ($child.NodeType -eq 'Element' -and $child.LocalName -eq 'p') {
            $t = Get-ParaText $child $nsm
            if ($t) { $title = $t; break }
        }
    }

    $caption = ''
    $heading = ''
    $prepItems = New-Object System.Collections.ArrayList
    $reqItems  = New-Object System.Collections.ArrayList
    $phases    = New-Object System.Collections.ArrayList
    $rxOptional = New-Object System.Text.RegularExpressions.Regex('\.?\s*$')
    $tableNo = 0
    $tableHeads = New-Object System.Collections.ArrayList

    foreach ($child in $body.ChildNodes) {
        if ($child.NodeType -ne 'Element') { continue }
        $tag = $child.LocalName

        if ($tag -eq 'p') {
            $style = Get-StyleName $child $nsm $styleMap
            $text = Get-ParaText $child $nsm
            if (-not $text) { continue }
            if ($style.ToLower().StartsWith('heading')) {
                $heading = $text
            } elseif ($style.ToLower() -eq 'table title') {
                $caption = [regex]::Replace($text, '^Table\s*[\d-]*\.?\s*', '')
            }
            continue
        }
        if ($tag -ne 'tbl') { continue }

        $rows = @($child.SelectNodes('w:tr', $nsm))
        $tableNo++
        if ($rows.Count -eq 0) { continue }
        $head = (Get-RowCells $rows[0] $nsm)
        $nonblank = @($head | Where-Object { $_ -ne '' })
        $seen = ('Table ' + $tableNo + ' (' + $rows.Count + ' rows): ' + (($head | ForEach-Object { if ($_.Length -gt 30) { $_.Substring(0, 30) } else { $_ } }) -join ' | '))
        [void]$tableHeads.Add($seen)

        # Execution plan tables
        if ($head.Count -ge 4 -and $head[0].ToLower() -eq 'action' -and $head[1].ToLower().StartsWith('impact')) {
            $name = ([regex]::Replace($caption, '\s*execution plan\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Trim()
            if (-not $name) { $name = $caption }
            if (-not $name) { $name = 'Execution' }
            $items = New-Object System.Collections.ArrayList
            for ($ri = 1; $ri -lt $rows.Count; $ri++) {
                $cells = (Get-RowCells $rows[$ri] $nsm)
                if ($cells.Count -eq 0 -or -not $cells[0]) { continue }
                $action   = [regex]::Replace($cells[0], '^\d+\.\s*', '')
                $impact   = ''; if ($cells.Count -gt 1) { $impact = $cells[1] }
                $required = ''; if ($cells.Count -gt 2) { $required = $cells[2] }
                $downtime = ''; if ($cells.Count -gt 3) { $downtime = $cells[3] }
                $kind = 'Task'
                if ($action.ToLower().StartsWith('validate')) { $kind = 'Checkpoint' }
                if ($required.ToLower().StartsWith('no')) { $action = $rxOptional.Replace($action, ' (optional)', 1) }
                $note = ($caption + '. Impact: ' + $impact + ' VM downtime: ' + $downtime).Trim()
                [void]$items.Add(@{ Kind = $kind; Text = $action; Note = $note })
            }
            if ($items.Count -gt 0) { [void]$phases.Add(@{ Name = $name; Items = $items }) }
            continue
        }

        # Preparation checklist: blank first column, "Description" second
        if ($nonblank.Count -eq 1 -and $nonblank[0] -ceq 'Description' -and $heading.ToLower().Contains('prepar')) {
            for ($ri = 1; $ri -lt $rows.Count; $ri++) {
                $cells = (Get-RowCells $rows[$ri] $nsm)
                if ($cells.Count -eq 0) { continue }
                $txt = $cells[$cells.Count - 1]
                if ($txt) { [void]$prepItems.Add(@{ Kind = 'Pre-check'; Text = (Strip-See $txt); Note = 'Convergence Preparation checklist' }) }
            }
            continue
        }

        # Requirements tables: "Item" + "...Requirement(s)"
        if ($head.Count -ge 2 -and $head[0].ToLower() -eq 'item' -and $head[1].ToLower().Contains('requirement')) {
            for ($ri = 1; $ri -lt $rows.Count; $ri++) {
                $tcs = @($rows[$ri].SelectNodes('w:tc', $nsm))
                if ($tcs.Count -lt 2) { continue }
                $label = Get-ParaText $tcs[0] $nsm
                $paras = (Get-CellParas $tcs[1] $nsm)
                if (-not $label -or $paras.Count -eq 0) { continue }
                if ($heading.ToLower().Contains('general')) { $src = 'General requirements: ' + $label } else { $src = 'Requirements table: ' + $label }
                if ($label.ToLower().StartsWith('minimum component')) {
                    [void]$reqItems.Add(@{ Kind = 'Pre-check'; Text = ($label + ': ' + ($paras -join '; ')); Note = $src })
                } else {
                    foreach ($para in $paras) {
                        [void]$reqItems.Add(@{ Kind = 'Pre-check'; Text = ($label + ': ' + $para); Note = $src })
                    }
                }
            }
        }
    }

    $prep = New-Object System.Collections.ArrayList
    foreach ($x in $prepItems) { [void]$prep.Add($x) }
    foreach ($x in $reqItems)  { [void]$prep.Add($x) }
    if ($prep.Count -gt 0) {
        [void]$prep.Add(@{ Kind = 'Checkpoint'; Text = 'Preparation complete: go or no-go decision recorded'; Note = 'Added by the script' })
        $phases.Insert(0, @{ Name = 'Preparation'; Items = $prep })
    }
    $hasExec = $false
    foreach ($ph in $phases) { if ($ph.Name -ne 'Preparation') { $hasExec = $true } }
    $mode = 'plan'
    if (-not $hasExec) {
        $proc = Read-ProcedurePlan $body $nsm $styleMap (Get-NumberFormats $numText)
        if ($proc.Count -eq 0) {
            $shown = @($tableHeads | Select-Object -First 15)
            $msg = 'No execution plan table found (a table headed Action / Impact to Environment / Required / VM Downtime), and no numbered steps under headings either.'
            $msg += [Environment]::NewLine + 'The document has ' + $tableNo + ' tables. Their first rows are:' + [Environment]::NewLine + ($shown -join [Environment]::NewLine)
            throw $msg
        }
        $phases = $proc
        $mode = 'procedure'
    }
    if (-not $title) { $title = [System.IO.Path]::GetFileNameWithoutExtension($path) }
    return [pscustomobject]@{ Title = $title; Phases = $phases; Mode = $mode }
}

# ---------------------------------------------------------------------------
# Writing the workbook (raw OOXML)
# ---------------------------------------------------------------------------

$script:NAVY   = '1F3864'
$script:INPUT  = 'FFF2CC'
$script:STEP_A = 'D9E2F3'
$script:STEP_B = 'FFE699'
$script:GREY   = 'BFBFBF'

function New-StyleBook {
    $sb = @{
        Fonts = New-Object System.Collections.ArrayList
        Fills = New-Object System.Collections.ArrayList
        Xfs   = New-Object System.Collections.ArrayList
        Dxfs  = New-Object System.Collections.ArrayList
        FKey  = @{}
        FlKey = @{}
        XKey  = @{}
    }
    [void]$sb.Fills.Add('<fill><patternFill patternType="none"/></fill>')
    [void]$sb.Fills.Add('<fill><patternFill patternType="gray125"/></fill>')
    # Font 0 and style 0 must be the plain default, because empty cells use them.
    [void](Get-StyleId $sb -sz 10)
    return $sb
}

function Get-FontId($sb, [int]$sz, [bool]$b, [bool]$i, $color, [string]$name) {
    $key = "$sz|$b|$i|$color|$name"
    if (-not $sb.FKey.ContainsKey($key)) {
        $x = '<font>'
        if ($b) { $x += '<b/>' }
        if ($i) { $x += '<i/>' }
        $x += '<sz val="' + $sz + '"/>'
        if ($color) { $x += '<color rgb="FF' + $color + '"/>' }
        $x += '<name val="' + $name + '"/></font>'
        [void]$sb.Fonts.Add($x)
        $sb.FKey[$key] = $sb.Fonts.Count - 1
    }
    return [int]$sb.FKey[$key]
}

function Get-FillId($sb, $rgb) {
    if (-not $rgb) { return 0 }
    if (-not $sb.FlKey.ContainsKey($rgb)) {
        [void]$sb.Fills.Add('<fill><patternFill patternType="solid"><fgColor rgb="FF' + $rgb + '"/><bgColor indexed="64"/></patternFill></fill>')
        $sb.FlKey[$rgb] = $sb.Fills.Count - 1
    }
    return [int]$sb.FlKey[$rgb]
}

function Get-StyleId($sb, [int]$sz = 10, [bool]$b = $false, [bool]$i = $false, $color = $null, $fill = $null,
                     [bool]$border = $false, [int]$nf = 0, [string]$h = '', [string]$v = '', [bool]$wrap = $false, [string]$name = 'Arial') {
    $fontId = Get-FontId $sb $sz $b $i $color $name
    $fillId = Get-FillId $sb $fill
    $borderId = 0; if ($border) { $borderId = 1 }
    $key = "$fontId|$fillId|$borderId|$nf|$h|$v|$wrap"
    if (-not $sb.XKey.ContainsKey($key)) {
        $al = ''
        if ($h -or $v -or $wrap) {
            $al = '<alignment'
            if ($h) { $al += ' horizontal="' + $h + '"' }
            if ($v) { $al += ' vertical="' + $v + '"' }
            if ($wrap) { $al += ' wrapText="1"' }
            $al += '/>'
        }
        $x = '<xf numFmtId="' + $nf + '" fontId="' + $fontId + '" fillId="' + $fillId + '" borderId="' + $borderId + '" xfId="0"'
        if ($nf -ne 0) { $x += ' applyNumberFormat="1"' }
        $x += ' applyFont="1"'
        if ($fillId -ne 0) { $x += ' applyFill="1"' }
        if ($borderId -ne 0) { $x += ' applyBorder="1"' }
        if ($al) { $x += ' applyAlignment="1">' + $al + '</xf>' } else { $x += '/>' }
        [void]$sb.Xfs.Add($x)
        $sb.XKey[$key] = $sb.Xfs.Count - 1
    }
    return [int]$sb.XKey[$key]
}

function Add-Dxf($sb, $color = $null, [bool]$bold = $false, $fill = $null, $size = $null) {
    $x = '<dxf>'
    if ($color -or $bold -or $size) {
        $x += '<font>'
        if ($bold) { $x += '<b/>' }
        if ($size) { $x += '<sz val="' + $size + '"/>' }
        if ($color) { $x += '<color rgb="FF' + $color + '"/>' }
        $x += '</font>'
    }
    if ($fill) { $x += '<fill><patternFill><bgColor rgb="FF' + $fill + '"/></patternFill></fill>' }
    $x += '</dxf>'
    [void]$sb.Dxfs.Add($x)
    return $sb.Dxfs.Count - 1
}

function Get-StylesXml($sb) {
    $g = $script:GREY
    $x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $x += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    $x += '<numFmts count="5"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy hh:mm"/><numFmt numFmtId="166" formatCode="0.0"/>'
    $x += '<numFmt numFmtId="167" formatCode="&quot;Window &quot;0"/><numFmt numFmtId="168" formatCode="ddd\ dd/mm/yyyy\ hh:mm"/></numFmts>'
    $x += '<fonts count="' + $sb.Fonts.Count + '">' + ($sb.Fonts -join '') + '</fonts>'
    $x += '<fills count="' + $sb.Fills.Count + '">' + ($sb.Fills -join '') + '</fills>'
    $x += '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    $x += '<border><left style="thin"><color rgb="FF' + $g + '"/></left><right style="thin"><color rgb="FF' + $g + '"/></right>'
    $x += '<top style="thin"><color rgb="FF' + $g + '"/></top><bottom style="thin"><color rgb="FF' + $g + '"/></bottom><diagonal/></border></borders>'
    $x += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    $x += '<cellXfs count="' + $sb.Xfs.Count + '">' + ($sb.Xfs -join '') + '</cellXfs>'
    $x += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    $x += '<dxfs count="' + $sb.Dxfs.Count + '">' + ($sb.Dxfs -join '') + '</dxfs></styleSheet>'
    return $x
}

function New-Sheet([string]$name) {
    return @{
        Name   = $name
        Rows   = New-Object 'System.Collections.Generic.SortedDictionary[int,object]'
        RowH   = @{}
        Merges = New-Object System.Collections.ArrayList
        Cols   = New-Object System.Collections.ArrayList
        Cf     = New-Object System.Collections.ArrayList
        Dv     = New-Object System.Collections.ArrayList
        Freeze = $null
        Grid   = $true
    }
}

# kind: s = text, n = number, f = formula (cached = value shown before Excel recalculates; ctype s = text result)
function Put-Cell($sh, [int]$r, [int]$c, $val, [int]$s, [string]$kind = 's', $cached = $null, [string]$ctype = 'n') {
    $ref = (Col-Name $c) + $r
    if ($null -eq $val -and $kind -ne 'f') {
        $x = '<c r="' + $ref + '" s="' + $s + '"/>'
    } elseif ($kind -eq 's') {
        $x = '<c r="' + $ref + '" s="' + $s + '" t="inlineStr"><is><t xml:space="preserve">' + (Esc-Xml ([string]$val)) + '</t></is></c>'
    } elseif ($kind -eq 'n') {
        $x = '<c r="' + $ref + '" s="' + $s + '"><v>' + ([string]$val) + '</v></c>'
    } else {
        $t = ''; if ($ctype -eq 's') { $t = ' t="str"' }
        $v = ''; if ($null -ne $cached) { $v = '<v>' + (Esc-Xml ([string]$cached)) + '</v>' }
        $x = '<c r="' + $ref + '" s="' + $s + '"' + $t + '><f>' + (Esc-Xml ([string]$val)) + '</f>' + $v + '</c>'
    }
    if (-not $sh.Rows.ContainsKey($r)) {
        $sh.Rows[$r] = New-Object 'System.Collections.Generic.SortedDictionary[int,string]'
    }
    $sh.Rows[$r][$c] = $x
}

function Add-Cf($sh, [string]$sqref, $rules) {
    [void]$sh.Cf.Add(@{ Sq = $sqref; Rules = $rules })
}

function Get-SheetXml($sh) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
    $view = '<sheetView workbookViewId="0"'
    if (-not $sh.Grid) { $view += ' showGridLines="0"' }
    $view += '>'
    if ($null -ne $sh.Freeze) {
        $col = [int]$sh.Freeze[0]; $row = [int]$sh.Freeze[1]
        $tl = (Col-Name ($col + 1)) + ($row + 1)
        if ($col -gt 0 -and $row -gt 0) {
            $view += '<pane xSplit="' + $col + '" ySplit="' + $row + '" topLeftCell="' + $tl + '" activePane="bottomRight" state="frozen"/><selection pane="bottomRight"/>'
        } else {
            $view += '<pane ySplit="' + $row + '" topLeftCell="' + $tl + '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/>'
        }
    }
    $view += '</sheetView>'
    [void]$sb.Append('<sheetViews>' + $view + '</sheetViews><sheetFormatPr defaultRowHeight="15"/>')
    [void]$sb.Append('<cols>')
    foreach ($cd in $sh.Cols) { [void]$sb.Append('<col min="' + $cd[0] + '" max="' + $cd[1] + '" width="' + $cd[2] + '" customWidth="1"/>') }
    [void]$sb.Append('</cols><sheetData>')
    foreach ($r in $sh.Rows.Keys) {
        $h = ''
        if ($sh.RowH.ContainsKey($r)) { $h = ' ht="' + $sh.RowH[$r] + '" customHeight="1"' }
        [void]$sb.Append('<row r="' + $r + '"' + $h + '>')
        [void]$sb.Append(($sh.Rows[$r].Values -join ''))
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($sh.Merges.Count -gt 0) {
        [void]$sb.Append('<mergeCells count="' + $sh.Merges.Count + '">')
        foreach ($m in $sh.Merges) { [void]$sb.Append('<mergeCell ref="' + $m + '"/>') }
        [void]$sb.Append('</mergeCells>')
    }
    $p = 0
    foreach ($cf in $sh.Cf) {
        [void]$sb.Append('<conditionalFormatting sqref="' + $cf.Sq + '">')
        foreach ($rule in $cf.Rules) {
            $p++
            [void]$sb.Append('<cfRule type="cellIs" dxfId="' + $rule[2] + '" priority="' + $p + '" operator="' + $rule[0] + '"><formula>' + (Esc-Xml $rule[1]) + '</formula></cfRule>')
        }
        [void]$sb.Append('</conditionalFormatting>')
    }
    if ($sh.Dv.Count -gt 0) {
        [void]$sb.Append('<dataValidations count="' + $sh.Dv.Count + '">')
        foreach ($d in $sh.Dv) {
            [void]$sb.Append('<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + $d[0] + '"><formula1>' + (Esc-Xml $d[1]) + '</formula1></dataValidation>')
        }
        [void]$sb.Append('</dataValidations>')
    }
    [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>')
    return $sb.ToString()
}

function Est-Height([string]$text, [int]$cpl) {
    $lines = 0
    foreach ($part in $text.Split("`n")) {
        $len = [math]::Max(1, $part.Length)
        $lines += [int][math]::Ceiling($len / $cpl)
    }
    if ($lines -lt 1) { $lines = 1 }
    if ($lines -eq 1) { return 15 }
    return [math]::Round(11.5 * $lines + 3.5, 2)
}

function Sub-Tok([string]$t, $map) {
    foreach ($k in $map.Keys) { $t = $t.Replace($k, [string]$map[$k]) }
    return $t
}

function Write-Zip([string]$path, $parts) {
    $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($p in $parts) {
            $e = $zip.CreateEntry($p.Name, [System.IO.Compression.CompressionLevel]::Optimal)
            $s = $e.Open()
            $b = $script:Utf8.GetBytes($p.Text)
            $s.Write($b, 0, $b.Length)
            $s.Dispose()
        }
    } finally {
        $zip.Dispose()
    }
}

function Get-EstBasis([string]$kind, [string]$text, [string]$mode) {
    # Picks which row of the Sizing tab an item's time estimate comes from. The user can change it per row.
    $t = $text.ToLower()
    if ($kind -eq 'Checkpoint') {
        if ($mode -eq 'procedure') { return 'Quick check' }
        return 'Health checkpoint'
    }
    if ($kind -eq 'Pre-check') { return 'Preparation task' }
    if ($t -match '^(download|upload)\b') { return 'Download or upload' }
    if ($mode -eq 'procedure') { return 'Config step' }
    if ($t -match 'esx' -and $t -match 'upgrade|remediat|update') { return 'ESX host upgrade' }
    if ($t -match 'vcenter' -and $t -match 'upgrade|update') { return 'vCenter upgrade' }
    if ($t -match 'edge' -and $t -match 'upgrade|update') { return 'NSX Edge upgrade' }
    if ($t -match 'nsx' -and $t -match 'upgrade|update') { return 'NSX Manager upgrade' }
    if ($t -match '^(deploy|install)\b') { return 'Appliance deploy' }
    if ($t -match '^(upgrade|update|patch)\b') { return 'Appliance upgrade' }
    return 'Config step'
}

function Build-Workbook([string]$title, $phases, [string]$sourceName, [string]$outPath, [string]$mode = 'plan') {
    $st = New-StyleBook
    $lcTitle = $title.ToLower()
    if ($lcTitle.Contains('converge')) { $tracker = 'Converge Tracker' }
    elseif ($lcTitle.Contains('upgrade')) { $tracker = 'Upgrade Tracker' }
    else { $tracker = 'Tracker' }
    $T = "'" + $tracker + "'"
    $FIRST = 12
    $nRows = 0; $itemsTotal = 0
    foreach ($ph in $phases) { $nRows += 1 + $ph.Items.Count; $itemsTotal += $ph.Items.Count }
    $LAST = $FIRST + $nRows + 24
    $ISS_LAST = 204
    $blk = [string][char]0x2588
    $lgt = [string][char]0x2591
    $today = Get-Date
    $todayText = $today.ToString('dd/MM/yyyy', $script:Inv)

    $tok = @{ '[F]' = $FIRST; '[L]' = $LAST; '[I]' = $ISS_LAST; '[T]' = $T }

    # Tracker columns: A Ref, B Type, C Item, D Est, E Owner, F Planned, G Status, H RAG, I Actual start, J Actual end,
    # K Actual hrs, L Notes, M Plan reference, N Estimate basis
    # O Est. start, P Est. finish, Q Window (suggested times, worked out from the Sizing tab)
    $NCOL = 17
    $WIN_HRS = 8.0; $WIN_RESERVE = 0.2; $WIN_USABLE = 6.4
    $CWN = "'Change Windows'"
    $CW_ROWS = 30

    # Time-estimate rules. Hours are starting values, not measured timings; the user replaces them on the Sizing tab.
    $unitsDefault = @(1, 4, 1, 1, 2)
    $rules = @(
        @{ N = 'Quick check';         V = @(0.25, 0, 0, 0, 0);     C = 'Read-only checks and confirmations' },
        @{ N = 'Preparation task';    V = @(0.5, 0, 0, 0, 0);      C = 'Pre-checks: DNS, IP addresses, access, capacity' },
        @{ N = 'Config step';         V = @(0.5, 0, 0, 0, 0);      C = 'Settings and console steps' },
        @{ N = 'Download or upload';  V = @(1, 0, 0, 0, 0);        C = 'Large downloads, depot and PAK uploads' },
        @{ N = 'Appliance deploy';    V = @(2, 0, 0, 0, 0);        C = 'OVA deployment and first configuration' },
        @{ N = 'Appliance upgrade';   V = @(2, 0, 0, 0, 0);        C = 'PAK or bundle upgrade of an appliance' },
        @{ N = 'vCenter upgrade';     V = @(0.5, 0, 0, 2.5, 0);    C = 'Per vCenter Server' },
        @{ N = 'NSX Manager upgrade'; V = @(3, 0, 0, 0, 0);        C = 'NSX Manager cluster' },
        @{ N = 'NSX Edge upgrade';    V = @(0.5, 0, 0, 0, 0.75);   C = 'Per Edge node' },
        @{ N = 'ESX host upgrade';    V = @(0.5, 0.75, 0.25, 0, 0); C = 'One host at a time, including evacuation' },
        @{ N = 'Health checkpoint';   V = @(0.5, 0, 0, 0, 0);      C = 'Validate health, safe stopping point' }
    )
    $ruleHours = @{}
    foreach ($ru in $rules) {
        $h = 0.0
        for ($u = 0; $u -lt 5; $u++) { $h += [double]$ru.V[$u] * [double]$unitsDefault[$u] }
        $ruleHours[$ru.N] = [math]::Round($h, 2)
    }

    $dGreen = Add-Dxf $st 'FFFFFF' $false '00B050'
    $dAmber = Add-Dxf $st '000000' $false 'FFC000'
    $dRed   = Add-Dxf $st 'FFFFFF' $false 'FF0000'
    $dComplete = Add-Dxf $st '00703C' $true
    $dBlocked  = Add-Dxf $st 'C00000' $true
    $dNotStarted = Add-Dxf $st '7F7F7F'
    $dInProg   = Add-Dxf $st '1F4E79' $false 'DDEBF7'
    $dAmberBig = Add-Dxf $st 'C55A11' $true $null 20
    $dRedBig   = Add-Dxf $st 'C00000' $true $null 20
    $rag = @( @('equal', '"Green"', $dGreen), @('equal', '"Amber"', $dAmber), @('equal', '"Red"', $dRed) )

    $dash = New-Sheet 'Dashboard'
    $trk  = New-Sheet $tracker
    $cw   = New-Sheet 'Change Windows'
    $siz  = New-Sheet 'Sizing'
    $iss  = New-Sheet 'Issues Log'
    $dash.Grid = $false

    $NAVY = $script:NAVY; $INPUT = $script:INPUT; $STEP_A = $script:STEP_A; $STEP_B = $script:STEP_B

    # ------------------------------------------------------------ Tracker
    $ws = $trk
    Put-Cell $ws 1 1 ($title + ': Execution Tracker') (Get-StyleId $st -sz 14 -b $true)
    if ($mode -eq 'procedure') { $tail = 'One row per top-level numbered step under each main heading; sub-steps are counted in the notes.' }
    else { $tail = 'Items marked (optional) have Required = No in the plan.' }
    Put-Cell $ws 2 1 ('Source: ' + $sourceName + ', generated ' + $todayText + '. ' + $phases.Count + ' phases, ' + $itemsTotal + ' items. ' + $tail) (Get-StyleId $st -i $true -color '595959')
    Put-Cell $ws 3 7 'Progress summary' (Get-StyleId $st -b $true)
    $labels = @('Change reference', 'Customer / environment', 'Change owner', 'Window start', 'Window end')
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $r = 4 + $i
        Put-Cell $ws $r 1 $labels[$i] (Get-StyleId $st -b $true)
        [void]$ws.Merges.Add('A' + $r + ':B' + $r)
        $nfx = 0; if ($labels[$i].StartsWith('Window')) { $nfx = 165 }
        Put-Cell $ws $r 3 $null (Get-StyleId $st -fill $INPUT -border $true -h 'left' -nf $nfx)
    }
    $summary = @(
        @(4,  'Green', 'COUNTIF($H$[F]:$H$[L],"Green")', '00B050', 'FFFFFF', 0),
        @(5,  'Amber', 'COUNTIF($H$[F]:$H$[L],"Amber")', 'FFC000', $null, 0),
        @(6,  'Red',   'COUNTIF($H$[F]:$H$[L],"Red")', 'FF0000', 'FFFFFF', 0),
        @(7,  'Items complete (excl. step rows)', 'COUNTIF($G$[F]:$G$[L],"Complete")-COUNTIFS($B$[F]:$B$[L],"Step",$G$[F]:$G$[L],"Complete")', $null, $null, 0),
        @(8,  'Items total (excl. step rows)', 'COUNTA($C$[F]:$C$[L])-COUNTIF($B$[F]:$B$[L],"Step")', $null, $null, $itemsTotal),
        @(9,  'Percent complete', 'IF(J8=0,0,J7/J8)', $null, $null, 0),
        @(10, 'Open issues', 'COUNTIF(''Issues Log''!$H$5:$H$[I],"Open")', $null, $null, 0)
    )
    foreach ($row in $summary) {
        $r = [int]$row[0]
        $isFilled = ($null -ne $row[3])
        Put-Cell $ws $r 7 $row[1] (Get-StyleId $st -b $isFilled -color $row[4] -fill $row[3])
        Put-Cell $ws $r 8 $null (Get-StyleId $st -fill $row[3])
        Put-Cell $ws $r 9 $null (Get-StyleId $st -fill $row[3])
        [void]$ws.Merges.Add('G' + $r + ':I' + $r)
        $nfx = 0; if ($r -eq 9) { $nfx = 9 }
        Put-Cell $ws $r 10 (Sub-Tok $row[2] $tok) (Get-StyleId $st -b $true -border $true -h 'left' -nf $nfx) 'f' $row[5]
    }
    [void]$ws.Merges.Add('A10:F10')
    Put-Cell $ws 10 1 'Yellow cells are for you to fill in. Pick Status and RAG from the dropdowns. Green = on plan, Amber = issue being managed, Red = blocked or needs a decision. Log each issue in the Issues Log tab using the item Ref. Est. (hrs) is worked out from the Sizing tab; type over it to override. Actual (hrs) comes from the actual start and end.' (Get-StyleId $st -i $true -color '595959' -wrap $true -v 'top')
    $ws.RowH[10] = 66
    [void]$ws.Merges.Add('O10:Q10')
    $fHint = 'IF(ISNUMBER($C$7),"","Enter Window start (C7) to see suggested start and finish times.")'
    $hintStyle = Get-StyleId $st -i $true -color 'C00000' -wrap $true -v 'top'
    Put-Cell $ws 10 15 $fHint $hintStyle 'f' 'Enter Window start (C7) to see suggested start and finish times.' 's'
    Put-Cell $ws 10 16 $null $hintStyle
    Put-Cell $ws 10 17 $null $hintStyle
    Put-Cell $ws 3 15 'Suggested times (columns O to Q)' (Get-StyleId $st -b $true)
    $keyNote = Get-StyleId $st -i $true -color '595959' -wrap $true -v 'top'
    Put-Cell $ws 4 15 'Tasks run in the order listed, starting at Window start (C7). Each task takes its Est. (hrs). A task that does not fit in the hours left in a window starts the next window. Window length, reserve and weekends are set on the Sizing tab. The Change Windows tab has the full key and the suggested windows.' $keyNote
    foreach ($kr2 in 4..8) { foreach ($kc in 15..17) { if (-not ($kr2 -eq 4 -and $kc -eq 15)) { Put-Cell $ws $kr2 $kc $null $keyNote } } }
    [void]$ws.Merges.Add('O4:Q8')
    $heads = @('Ref', 'Type', 'Item', 'Est. (hrs)', 'Owner', 'Planned date', 'Status', 'RAG', 'Actual start', 'Actual end', 'Actual (hrs)', 'Notes / issues observed', 'Plan reference', 'Estimate basis', 'Est. start', 'Est. finish', 'Window')
    for ($c = 1; $c -le $heads.Count; $c++) {
        Put-Cell $ws 11 $c $heads[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center' -wrap $true)
    }
    $ws.RowH[11] = 24
    $ctr = @(1, 2, 4, 6, 7, 8, 9, 10, 11)
    $fEst = 'IF($N[r]="","",IFERROR(SUMPRODUCT(INDEX(Sizing!$B$14:$F$25,MATCH($N[r],Sizing!$A$14:$A$25,0),0),Sizing!$B$13:$F$13),"?"))'
    $fAct = 'IF(AND(ISNUMBER(I[r]),ISNUMBER(J[r]),J[r]>=I[r]),ROUND((J[r]-I[r])*24,2),"")'

    # Suggested times. Window number: tasks fill a window in order; a task that will not fit in the hours left starts the next window.
    $wPrev = 'MAX(1,MAX(Q$[h]:Q[q]))'
    $wUsed = 'SUMIFS($D$[h]:$D[q],$Q$[h]:$Q[q],' + $wPrev + ')'
    $fWin = 'IF(OR($B[r]="Step",NOT(ISNUMBER($D[r]))),"",' + $wPrev + '+IF(AND(' + $wUsed + '>0,ROUND(' + $wUsed + '+$D[r],4)>ROUND(Sizing!$F$8,4)),1,0))'
    $fWinStart = 'IF(Sizing!$F$6="Yes",WORKDAY(INT($C$7),$Q[r]-1),INT($C$7)+$Q[r]-1)+MOD($C$7,1)'
    $fStart = 'IF(OR($Q[r]="",NOT(ISNUMBER($C$7))),"",' + $fWinStart + '+SUMIFS($D$[h]:$D[q],$Q$[h]:$Q[q],$Q[r])/24)'
    $fFinish = 'IF($O[r]="","",$O[r]+$D[r]/24)'
    $winNo = 1; $winUsedSim = 0.0
    $winHrs = New-Object System.Collections.ArrayList
    [void]$winHrs.Add(0.0)

    $phEst = New-Object System.Collections.ArrayList
    $rowInfo = @{}
    $r = $FIRST
    for ($pi = 1; $pi -le $phases.Count; $pi++) {
        $ph = $phases[$pi - 1]
        $vals = @{ 1 = ('S' + $pi); 2 = 'Step'; 3 = ('Phase ' + $pi + ': ' + $ph.Name); 7 = 'Not started' }
        for ($c = 1; $c -le $NCOL; $c++) {
            $fillc = $STEP_B; if ($c -in 1, 2, 3, 13) { $fillc = $STEP_A }
            $hh = ''; if ($c -in 1, 2, 7) { $hh = 'center' }
            $sid = Get-StyleId $st -b $true -fill $fillc -border $true -h $hh -v 'center' -wrap $true
            Put-Cell $ws $r $c $vals[$c] $sid
        }
        $ws.RowH[$r] = 15
        $rowInfo[$r] = 'step'
        $r++
        $sumEst = 0.0
        for ($ii = 1; $ii -le $ph.Items.Count; $ii++) {
            $it = $ph.Items[$ii - 1]
            $basis = Get-EstBasis $it.Kind $it.Text $mode
            $cachedEst = $ruleHours[$basis]
            $sumEst += [double]$cachedEst
            if (($winUsedSim -gt 0) -and ([math]::Round($winUsedSim + [double]$cachedEst, 4) -gt [math]::Round($WIN_USABLE, 4))) {
                $winNo++; $winUsedSim = 0.0; [void]$winHrs.Add(0.0)
            }
            $winUsedSim += [double]$cachedEst
            $winHrs[$winNo - 1] = [math]::Round([double]$winHrs[$winNo - 1] + [double]$cachedEst, 2)
            $rowInfo[$r] = $winNo
            $rm = @{ '[r]' = $r }
            for ($c = 1; $c -le $NCOL; $c++) {
                $nfx = 0
                if ($c -eq 6) { $nfx = 164 } elseif ($c -in 9, 10) { $nfx = 165 } elseif ($c -in 4, 11) { $nfx = 166 }
                $fillc = $null; if ($c -in 4, 5, 6, 7, 8, 9, 10, 12, 14) { $fillc = $INPUT }
                $hh = ''; if ($ctr -contains $c) { $hh = 'center' }
                $sid = Get-StyleId $st -fill $fillc -border $true -nf $nfx -h $hh -v 'center' -wrap $true
                if ($c -eq 1) { Put-Cell $ws $r $c ([string]$pi + '.' + $ii) $sid }
                elseif ($c -eq 2) { Put-Cell $ws $r $c $it.Kind $sid }
                elseif ($c -eq 3) { Put-Cell $ws $r $c $it.Text $sid }
                elseif ($c -eq 4) { Put-Cell $ws $r $c (Sub-Tok $fEst $rm) $sid 'f' $cachedEst }
                elseif ($c -eq 7) { Put-Cell $ws $r $c 'Not started' $sid }
                elseif ($c -eq 11) { Put-Cell $ws $r $c (Sub-Tok $fAct $rm) $sid 'f' '' 's' }
                elseif ($c -eq 12) { Put-Cell $ws $r $c $null $sid }
                elseif ($c -eq 13) { Put-Cell $ws $r $c $it.Note $sid }
                elseif ($c -eq 14) { Put-Cell $ws $r $c $basis $sid }
                else { Put-Cell $ws $r $c $null $sid }
            }
            $ws.RowH[$r] = [math]::Max([double](Est-Height $it.Text 75), [double](Est-Height $it.Note 36))
            $r++
        }
        [void]$phEst.Add([math]::Round($sumEst, 2))
    }
    for ($rr = $r; $rr -le $LAST; $rr++) {
        $rm = @{ '[r]' = $rr }
        for ($c = 1; $c -le $NCOL; $c++) {
            $nfx = 0
            if ($c -eq 6) { $nfx = 164 } elseif ($c -in 9, 10) { $nfx = 165 } elseif ($c -in 4, 11) { $nfx = 166 }
            $hh = ''; if ($ctr -contains $c) { $hh = 'center' }
            $sid = Get-StyleId $st -fill $INPUT -border $true -nf $nfx -h $hh -v 'center' -wrap $true
            if ($c -eq 11) { Put-Cell $ws $rr $c (Sub-Tok $fAct $rm) $sid 'f' '' 's' }
            else { Put-Cell $ws $rr $c $null $sid }
        }
    }
    # Suggested start, finish and window on every row (step and blank rows return an empty string)
    for ($rr = $FIRST; $rr -le $LAST; $rr++) {
        $rm = @{ '[r]' = $rr; '[q]' = ($rr - 1); '[h]' = ($FIRST - 1) }
        $isStep = ($rowInfo.ContainsKey($rr) -and ($rowInfo[$rr] -is [string]))
        $fillw = $null; if ($isStep) { $fillw = $STEP_B }
        $bw = $isStep
        $styQ = Get-StyleId $st -b $bw -fill $fillw -border $true -nf 167 -h 'center' -v 'center'
        $styT = Get-StyleId $st -b $bw -fill $fillw -border $true -nf 168 -h 'center' -v 'center'
        if ($rowInfo.ContainsKey($rr) -and ($rowInfo[$rr] -is [int])) {
            Put-Cell $ws $rr 17 (Sub-Tok $fWin $rm) $styQ 'f' $rowInfo[$rr]
        } else {
            Put-Cell $ws $rr 17 (Sub-Tok $fWin $rm) $styQ 'f' '' 's'
        }
        Put-Cell $ws $rr 15 (Sub-Tok $fStart $rm) $styT 'f' '' 's'
        Put-Cell $ws $rr 16 (Sub-Tok $fFinish $rm) $styT 'f' '' 's'
    }
    foreach ($cd in @(@(1,1,7), @(2,2,16), @(3,3,70), @(4,4,10), @(5,5,14), @(6,6,13), @(7,7,14), @(8,8,10), @(9,10,15), @(11,11,11), @(12,12,55), @(13,13,36), @(14,14,22), @(15,16,21), @(17,17,11))) { [void]$ws.Cols.Add($cd) }
    $ws.Freeze = @(3, 11)
    [void]$ws.Dv.Add(@(('G' + $FIRST + ':G' + $LAST), '"Not started,In progress,Complete,Blocked,Rolled back"'))
    [void]$ws.Dv.Add(@(('H' + $FIRST + ':H' + $LAST), '"Green,Amber,Red"'))
    [void]$ws.Dv.Add(@(('N' + $FIRST + ':N' + $LAST), 'Sizing!$A$14:$A$25'))
    Add-Cf $ws ('G' + $FIRST + ':G' + $LAST) @( @('equal', '"Complete"', $dComplete), @('equal', '"Blocked"', $dBlocked), @('equal', '"Rolled back"', $dBlocked) )
    Add-Cf $ws ('H' + $FIRST + ':H' + $LAST) $rag

    # ------------------------------------------------------------ Dashboard
    $ws = $dash
    $totalEst = 0.0
    foreach ($e in $phEst) { $totalEst += [double]$e }
    $totalEst = [math]::Round($totalEst, 2)
    $lastRow = 8 + $phases.Count
    $totRow = $lastRow + 1
    Put-Cell $ws 1 1 ($title + ': Status at a glance') (Get-StyleId $st -sz 16 -b $true)
    $ws.RowH[1] = 20.65
    Put-Cell $ws 2 1 ('Updates itself from the ' + $tracker + ', Sizing and Issues Log tabs. Nothing to type on this page.') (Get-StyleId $st -i $true -color '595959')
    $tiles = @(
        @(1, 2,   'Overall complete', ($T + '!J9'),  9, 0),
        @(3, 4,   'Items done',       ($T + '!J7'),  0, 0),
        @(5, 6,   'Items total',      ($T + '!J8'),  0, $itemsTotal),
        @(7, 8,   'Open issues',      ($T + '!J10'), 0, 0),
        @(9, 10,  'Red items',        ($T + '!J6'),  0, 0),
        @(11, 12, 'Est. hours left',  ('L' + $totRow), 166, $totalEst)
    )
    foreach ($tile in $tiles) {
        $a = [int]$tile[0]; $b = [int]$tile[1]
        foreach ($c in @($a, $b)) {
            $lab = $null; if ($c -eq $a) { $lab = $tile[2] }
            Put-Cell $ws 4 $c $lab (Get-StyleId $st -sz 9 -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center')
            $sid = Get-StyleId $st -sz 20 -b $true -border $true -h 'center' -v 'center' -nf ([int]$tile[4])
            if ($c -eq $a) { Put-Cell $ws 5 $c $tile[3] $sid 'f' $tile[5] } else { Put-Cell $ws 5 $c $null $sid }
        }
        [void]$ws.Merges.Add((Col-Name $a) + '4:' + (Col-Name $b) + '4')
        [void]$ws.Merges.Add((Col-Name $a) + '5:' + (Col-Name $b) + '5')
    }
    $ws.RowH[5] = 33.75
    Add-Cf $ws 'G5' @( ,@('greaterThan', '0', $dAmberBig) )
    Add-Cf $ws 'I5' @( ,@('greaterThan', '0', $dRedBig) )
    $dh = @('#', 'Phase', 'Status', 'RAG', 'Done', 'Total', '% done', 'Progress', 'Remaining', 'Est. hrs', 'Actual hrs', 'Hrs left')
    for ($c = 1; $c -le $dh.Count; $c++) {
        Put-Cell $ws 8 $c $dh[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center')
    }
    $ws.RowH[8] = 19.5
    $colA = $T + '!$A$' + $FIRST + ':$A$' + $LAST
    $colB = $T + '!$B$' + $FIRST + ':$B$' + $LAST
    $colE = $T + '!$D$' + $FIRST + ':$D$' + $LAST
    $colS = $T + '!$G$' + $FIRST + ':$G$' + $LAST
    $colR = $T + '!$H$' + $FIRST + ':$H$' + $LAST
    $colK = $T + '!$K$' + $FIRST + ':$K$' + $LAST
    $fStatus = 'IF(COUNTIFS([A],$A[r]&".*",[S],"Blocked")+COUNTIFS([A],$A[r]&".*",[S],"Rolled back")>0,"Blocked",IF(E[r]=F[r],"Complete",IF(OR(E[r]>0,COUNTIFS([A],$A[r]&".*",[S],"In progress")>0),"In progress","Not started")))'
    $fRag    = 'IF(OR(C[r]="Blocked",COUNTIFS([A],$A[r]&".*",[Rg],"Red")>0),"Red",IF(COUNTIFS([A],$A[r]&".*",[Rg],"Amber")>0,"Amber",IF(C[r]="Not started","","Green")))'
    $fDone   = 'COUNTIFS([A],$A[r]&".*",[S],"Complete")'
    $fTotal  = 'COUNTIFS([A],$A[r]&".*",[B],"<>Step")'
    $fBar    = 'REPT("' + $blk + '",ROUND(G[r]*20,0))&REPT("' + $lgt + '",20-ROUND(G[r]*20,0))'
    $fEstP   = 'SUMIFS([E],[A],$A[r]&".*",[B],"<>Step")'
    $fActP   = 'SUMIFS([K],[A],$A[r]&".*",[B],"<>Step")'
    $fLeftP  = 'SUMIFS([E],[A],$A[r]&".*",[B],"<>Step",[S],"<>Complete")'
    $barCached = $lgt * 20
    for ($i = 1; $i -le $phases.Count; $i++) {
        $ph = $phases[$i - 1]
        $r = 8 + $i
        $m = @{ '[A]' = $colA; '[B]' = $colB; '[E]' = $colE; '[S]' = $colS; '[Rg]' = $colR; '[K]' = $colK; '[r]' = $r }
        $ws.RowH[$r] = 19.5
        $nItems = $ph.Items.Count
        $base = Get-StyleId $st -border $true -h 'center' -v 'center'
        $hrsStyle = Get-StyleId $st -border $true -h 'center' -v 'center' -nf 166
        Put-Cell $ws $r 1 $i $base 'n'
        Put-Cell $ws $r 2 $ph.Name (Get-StyleId $st -border $true -h 'left' -v 'center')
        Put-Cell $ws $r 3 (Sub-Tok $fStatus $m) $base 'f' 'Not started' 's'
        Put-Cell $ws $r 4 (Sub-Tok $fRag $m) $base 'f' '' 's'
        Put-Cell $ws $r 5 (Sub-Tok $fDone $m) $base 'f' 0
        Put-Cell $ws $r 6 (Sub-Tok $fTotal $m) $base 'f' $nItems
        Put-Cell $ws $r 7 (Sub-Tok 'IF(F[r]=0,0,E[r]/F[r])' $m) (Get-StyleId $st -border $true -h 'center' -v 'center' -nf 9) 'f' 0
        Put-Cell $ws $r 8 (Sub-Tok $fBar $m) (Get-StyleId $st -color '2E75B6' -border $true -h 'left' -v 'center' -name 'Consolas') 'f' $barCached 's'
        Put-Cell $ws $r 9 (Sub-Tok 'F[r]-E[r]' $m) $base 'f' $nItems
        Put-Cell $ws $r 10 (Sub-Tok $fEstP $m) $hrsStyle 'f' $phEst[$i - 1]
        Put-Cell $ws $r 11 (Sub-Tok $fActP $m) $hrsStyle 'f' 0
        Put-Cell $ws $r 12 (Sub-Tok $fLeftP $m) $hrsStyle 'f' $phEst[$i - 1]
    }
    # Totals row
    $totStyle = Get-StyleId $st -b $true -fill $STEP_A -border $true -h 'center' -v 'center'
    $totHrs = Get-StyleId $st -b $true -fill $STEP_A -border $true -h 'center' -v 'center' -nf 166
    Put-Cell $ws $totRow 1 $null $totStyle
    Put-Cell $ws $totRow 2 'Total' (Get-StyleId $st -b $true -fill $STEP_A -border $true -h 'left' -v 'center')
    for ($c = 3; $c -le 9; $c++) { Put-Cell $ws $totRow $c $null $totStyle }
    Put-Cell $ws $totRow 10 ('SUM(J9:J' + $lastRow + ')') $totHrs 'f' $totalEst
    Put-Cell $ws $totRow 11 ('SUM(K9:K' + $lastRow + ')') $totHrs 'f' 0
    Put-Cell $ws $totRow 12 ('SUM(L9:L' + $lastRow + ')') $totHrs 'f' $totalEst
    $ws.RowH[$totRow] = 19.5
    $fCal = 'IF(SUMIFS([E],[S],"Complete",[K],">0")=0,"Calibration: no completed items with actual times yet.","Calibration: completed items with actual times took "&TEXT(SUMIFS([K],[S],"Complete",[K],">0")/SUMIFS([E],[S],"Complete",[K],">0"),"0%")&" of their estimate so far. Adjust the Sizing tab hours for the next environment.")'
    $mc = @{ '[E]' = $colE; '[S]' = $colS; '[K]' = $colK }
    Put-Cell $ws ($totRow + 1) 2 (Sub-Tok $fCal $mc) (Get-StyleId $st -sz 9 -i $true -color '595959') 'f' 'Calibration: no completed items with actual times yet.' 's'
    Put-Cell $ws ($totRow + 2) 2 'RAG per phase: Red if any item is Red or Blocked, Amber if any is Amber, otherwise Green once the phase has started.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    Put-Cell $ws ($totRow + 3) 2 'Hours are estimates from the Sizing tab (starting values, not measured timings) and are not a commitment.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    $fSug = 'IF(' + $CWN + '!$C$5=0,"","Suggested change windows: "&' + $CWN + '!$C$5&IF(ISNUMBER(' + $CWN + '!$C$7)," (from "&TEXT(' + $T + '!$C$7,"ddd dd mmm yyyy hh:mm")&" to "&TEXT(' + $CWN + '!$C$7,"ddd dd mmm yyyy hh:mm")&")","")&". See the Change Windows tab.")'
    Put-Cell $ws ($totRow + 4) 2 $fSug (Get-StyleId $st -sz 9 -i $true -color '595959') 'f' ('Suggested change windows: ' + $winNo + '. See the Change Windows tab.') 's'
    Add-Cf $ws ('C9:C' + $lastRow) @( @('equal', '"Complete"', $dComplete), @('equal', '"In progress"', $dInProg), @('equal', '"Blocked"', $dBlocked), @('equal', '"Not started"', $dNotStarted) )
    Add-Cf $ws ('D9:D' + $lastRow) $rag
    foreach ($cd in @(@(1,1,5), @(2,2,42), @(3,3,14), @(4,4,10), @(5,6,8), @(7,7,9), @(8,8,26), @(9,9,11), @(10,12,10))) { [void]$ws.Cols.Add($cd) }

    # ------------------------------------------------------------ Sizing
    $ws = $siz
    Put-Cell $ws 1 1 'Sizing and time estimates' (Get-StyleId $st -sz 14 -b $true)
    Put-Cell $ws 2 1 'Enter the size of this environment in the yellow cells. Each tracker row looks up its Estimate basis in the table below: hours = fixed hours + (hours per unit x number of units).' (Get-StyleId $st -i $true -color '595959')
    Put-Cell $ws 3 1 'The hours below are starting values, not measured timings. Replace them with your own, and use the calibration line on the Dashboard after the first environment.' (Get-StyleId $st -i $true -color '595959')
    $inLabels = @('ESX hosts', 'Clusters', 'vCenter Servers', 'NSX Edge nodes')
    $inCached = @(4, 1, 1, 2)
    for ($i = 0; $i -lt 4; $i++) {
        Put-Cell $ws (4 + $i) 1 $inLabels[$i] (Get-StyleId $st -b $true -border $true)
        Put-Cell $ws (4 + $i) 2 $inCached[$i] (Get-StyleId $st -fill $INPUT -border $true -h 'center') 'n'
    }
    Put-Cell $ws 9 1 'Example values are filled in. Replace them with the numbers for this environment.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    # Change window settings, used for the suggested start and finish times on the tracker
    $setLab = Get-StyleId $st -b $true -border $true
    $setIn  = Get-StyleId $st -fill $INPUT -border $true -h 'center'
    $setInP = Get-StyleId $st -fill $INPUT -border $true -h 'center' -nf 9
    $setCalc = Get-StyleId $st -border $true -h 'center' -nf 166
    $setNames = @('Hours per window', 'Reserve at end of window', 'Skip weekends', 'Window length used (hrs)', 'Usable hours per window')
    for ($i = 0; $i -lt 5; $i++) {
        Put-Cell $ws (4 + $i) 4 $setNames[$i] $setLab
        Put-Cell $ws (4 + $i) 5 $null $setLab
        [void]$ws.Merges.Add('D' + (4 + $i) + ':E' + (4 + $i))
    }
    Put-Cell $ws 4 6 $WIN_HRS $setIn 'n'
    Put-Cell $ws 5 6 $WIN_RESERVE $setInP 'n'
    Put-Cell $ws 6 6 'Yes' $setIn
    $fLen = 'IF(AND(ISNUMBER(' + $T + '!$C$7),ISNUMBER(' + $T + '!$C$8),' + $T + '!$C$8>' + $T + '!$C$7),(' + $T + '!$C$8-' + $T + '!$C$7)*24,F4)'
    Put-Cell $ws 7 6 $fLen $setCalc 'f' $WIN_HRS
    Put-Cell $ws 8 6 'F7*(1-F5)' $setCalc 'f' $WIN_USABLE
    [void]$ws.Dv.Add(@('F6', '"Yes,No"'))
    Put-Cell $ws 10 1 'Window length comes from Window start and Window end on the tracker when both are filled in, otherwise from Hours per window.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    Put-Cell $ws 11 1 'The reserve is kept free at the end of each window for overruns and validation.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    $sh = @('Estimate basis', 'Fixed hrs', 'Per ESX host', 'Per cluster', 'Per vCenter', 'Per Edge node', 'What it covers')
    for ($c = 1; $c -le $sh.Count; $c++) {
        Put-Cell $ws 12 $c $sh[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center' -wrap $true)
    }
    $ws.RowH[12] = 24
    $unitStyle = Get-StyleId $st -b $true -fill $STEP_A -border $true -h 'center'
    Put-Cell $ws 13 1 'Units in this environment' (Get-StyleId $st -b $true -fill $STEP_A -border $true)
    Put-Cell $ws 13 2 1 $unitStyle 'n'
    Put-Cell $ws 13 3 '$B$4' $unitStyle 'f' 4
    Put-Cell $ws 13 4 '$B$5' $unitStyle 'f' 1
    Put-Cell $ws 13 5 '$B$6' $unitStyle 'f' 1
    Put-Cell $ws 13 6 '$B$7' $unitStyle 'f' 2
    Put-Cell $ws 13 7 $null (Get-StyleId $st -fill $STEP_A -border $true)
    for ($k = 0; $k -lt 12; $k++) {
        $rr = 14 + $k
        $ru = $null; if ($k -lt $rules.Count) { $ru = $rules[$k] }
        $sidT = Get-StyleId $st -fill $INPUT -border $true
        $sidN = Get-StyleId $st -fill $INPUT -border $true -h 'center'
        if ($null -ne $ru) {
            Put-Cell $ws $rr 1 $ru.N $sidT
            for ($u = 0; $u -lt 5; $u++) { Put-Cell $ws $rr (2 + $u) ([double]$ru.V[$u]) $sidN 'n' }
            Put-Cell $ws $rr 7 $ru.C $sidT
        } else {
            for ($c = 1; $c -le 7; $c++) { Put-Cell $ws $rr $c $null $sidT }
        }
    }
    Put-Cell $ws 27 1 'Add your own estimate bases in the empty yellow rows; they appear in the Estimate basis dropdown on the tracker.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    foreach ($cd in @(@(1,1,30), @(2,6,14), @(7,7,52))) { [void]$ws.Cols.Add($cd) }

    # ------------------------------------------------------------ Change Windows
    $ws = $cw
    $rQ = $T + '!$Q$' + $FIRST + ':$Q$' + $LAST
    $rD = $T + '!$D$' + $FIRST + ':$D$' + $LAST
    $rA = $T + '!$A$' + $FIRST + ':$A$' + $LAST
    $rB = $T + '!$B$' + $FIRST + ':$B$' + $LAST
    $rP = $T + '!$P$' + $FIRST + ':$P$' + $LAST
    $grey = '595959'
    Put-Cell $ws 1 1 'Suggested change windows' (Get-StyleId $st -sz 14 -b $true)
    Put-Cell $ws 2 1 ('Worked out from the order of tasks on the ' + $tracker + ' tab and their Est. (hrs). Fill in Window start on the tracker (cell C7) and the start and end times appear. Nothing to type on this page.') (Get-StyleId $st -i $true -color $grey)
    $sumLab = Get-StyleId $st -b $true -border $true
    $sumVal = Get-StyleId $st -b $true -border $true -h 'left'
    $sumDt  = Get-StyleId $st -b $true -border $true -h 'left' -nf 168
    $sumHr  = Get-StyleId $st -b $true -border $true -h 'left' -nf 166
    $fFirst = 'IF(ISNUMBER(' + $T + '!$C$7),' + $T + '!$C$7,"Enter Window start on the tracker (C7)")'
    $fCount = 'IF(COUNT(' + $rQ + ')=0,0,MAX(' + $rQ + '))'
    $fTotal2 = 'SUMIFS(' + $rD + ',' + $rB + ',"<>Step")'
    $fLast = 'IF(COUNT(' + $rP + ')=0,"",MAX(' + $rP + '))'
    $fLong = 'COUNTIF(' + $rD + ',">"&Sizing!$F$8)'
    $sumRows = @(
        @(4, 'First window starts', $fFirst, $sumDt, 'Enter Window start on the tracker (C7)', 's'),
        @(5, 'Windows needed', $fCount, $sumVal, $winNo, 'n'),
        @(6, 'Total estimated hours', $fTotal2, $sumHr, $totalEst, 'n'),
        @(7, 'Last task finishes', $fLast, $sumDt, '', 's'),
        @(8, 'Tasks longer than one window', $fLong, $sumVal, 0, 'n')
    )
    foreach ($sr in $sumRows) {
        $r = [int]$sr[0]
        Put-Cell $ws $r 1 $sr[1] $sumLab
        Put-Cell $ws $r 2 $null $sumLab
        [void]$ws.Merges.Add('A' + $r + ':B' + $r)
        Put-Cell $ws $r 3 $sr[2] $sr[3] 'f' $sr[4] $sr[5]
    }
    $wh = @('Window', 'Suggested start', 'Suggested end', 'Tasks finish by', 'Est. hrs', 'Usable hrs', 'Spare hrs', 'First task', 'Last task', 'Wording for the customer')
    for ($c = 1; $c -le $wh.Count; $c++) {
        Put-Cell $ws 10 $c $wh[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center' -wrap $true)
    }
    $ws.RowH[10] = 24
    $cwBase = Get-StyleId $st -border $true -h 'center' -v 'center'
    $cwDt = Get-StyleId $st -border $true -h 'center' -v 'center' -nf 168
    $cwHr = Get-StyleId $st -border $true -h 'center' -v 'center' -nf 166
    $cwWin = Get-StyleId $st -b $true -border $true -h 'center' -v 'center' -nf 167
    $cwTxt = Get-StyleId $st -border $true -h 'left' -v 'center' -wrap $true
    $fWinA = 'IF([n]<=$C$5,[n],"")'
    $fWS = 'IF(OR($A[r]="",NOT(ISNUMBER(' + $T + '!$C$7))),"",IF(Sizing!$F$6="Yes",WORKDAY(INT(' + $T + '!$C$7),$A[r]-1),INT(' + $T + '!$C$7)+$A[r]-1)+MOD(' + $T + '!$C$7,1))'
    $fWE = 'IF($B[r]="","",$B[r]+Sizing!$F$7/24)'
    $fWF = 'IF($B[r]="","",$B[r]+$E[r]/24)'
    $fWH = 'IF($A[r]="","",SUMIFS(' + $rD + ',' + $rQ + ',$A[r]))'
    $fWU = 'IF($A[r]="","",Sizing!$F$8)'
    $fWG = 'IF($A[r]="","",ROUND($F[r]-$E[r],2))'
    $fW1 = 'IF($A[r]="","",INDEX(' + $rA + ',MATCH($A[r],' + $rQ + ',0)))'
    $fW2 = 'IF($A[r]="","",INDEX(' + $rA + ',SUMPRODUCT(MAX((' + $rQ + '=$A[r])*(ROW(' + $rQ + ')-' + $FIRST + '+1)))))'
    $fWT = 'IF($B[r]="","",TEXT($B[r],"ddd dd mmm yyyy hh:mm")&" to "&TEXT($C[r],"ddd dd mmm yyyy hh:mm")&" (tasks "&$H[r]&" to "&$I[r]&", about "&TEXT($E[r],"0.0")&" hrs)")'
    $firstRefs = New-Object System.Collections.ArrayList
    $lastRefs = New-Object System.Collections.ArrayList
    for ($n = 1; $n -le $winNo; $n++) { [void]$firstRefs.Add(''); [void]$lastRefs.Add('') }
    $simW = 1; $simU = 0.0
    $pi2 = 0
    foreach ($ph in $phases) {
        $pi2++; $ii2 = 0
        foreach ($it in $ph.Items) {
            $ii2++
            $hh2 = [double]$ruleHours[(Get-EstBasis $it.Kind $it.Text $mode)]
            if (($simU -gt 0) -and ([math]::Round($simU + $hh2, 4) -gt [math]::Round($WIN_USABLE, 4))) { $simW++; $simU = 0.0 }
            $simU += $hh2
            if ($firstRefs[$simW - 1] -eq '') { $firstRefs[$simW - 1] = [string]$pi2 + '.' + $ii2 }
            $lastRefs[$simW - 1] = [string]$pi2 + '.' + $ii2
        }
    }
    for ($n = 1; $n -le $CW_ROWS; $n++) {
        $r = 10 + $n
        $m = @{ '[r]' = $r; '[n]' = $n }
        $has = ($n -le $winNo)
        $cA = ''; $cT = 's'; if ($has) { $cA = $n; $cT = 'n' }
        Put-Cell $ws $r 1 (Sub-Tok $fWinA $m) $cwWin 'f' $cA $cT
        Put-Cell $ws $r 2 (Sub-Tok $fWS $m) $cwDt 'f' '' 's'
        Put-Cell $ws $r 3 (Sub-Tok $fWE $m) $cwDt 'f' '' 's'
        Put-Cell $ws $r 4 (Sub-Tok $fWF $m) $cwDt 'f' '' 's'
        if ($has) {
            Put-Cell $ws $r 5 (Sub-Tok $fWH $m) $cwHr 'f' $winHrs[$n - 1]
            Put-Cell $ws $r 6 (Sub-Tok $fWU $m) $cwHr 'f' $WIN_USABLE
            Put-Cell $ws $r 7 (Sub-Tok $fWG $m) $cwHr 'f' ([math]::Round($WIN_USABLE - [double]$winHrs[$n - 1], 2))
            Put-Cell $ws $r 8 (Sub-Tok $fW1 $m) $cwBase 'f' $firstRefs[$n - 1] 's'
            Put-Cell $ws $r 9 (Sub-Tok $fW2 $m) $cwBase 'f' $lastRefs[$n - 1] 's'
        } else {
            Put-Cell $ws $r 5 (Sub-Tok $fWH $m) $cwHr 'f' '' 's'
            Put-Cell $ws $r 6 (Sub-Tok $fWU $m) $cwHr 'f' '' 's'
            Put-Cell $ws $r 7 (Sub-Tok $fWG $m) $cwHr 'f' '' 's'
            Put-Cell $ws $r 8 (Sub-Tok $fW1 $m) $cwBase 'f' '' 's'
            Put-Cell $ws $r 9 (Sub-Tok $fW2 $m) $cwBase 'f' '' 's'
        }
        Put-Cell $ws $r 10 (Sub-Tok $fWT $m) $cwTxt 'f' '' 's'
    }
    $dNeg = Add-Dxf $st 'C00000' $true
    Add-Cf $ws ('G11:G' + (10 + $CW_ROWS)) @( ,@('lessThan', '0', $dNeg) )
    $kr = 11 + $CW_ROWS + 1
    $keyHead = Get-StyleId $st -b $true
    $keyTxt = Get-StyleId $st -color '404040' -wrap $true -v 'top'
    $keyI = Get-StyleId $st -sz 9 -i $true -color $grey -wrap $true -v 'top'
    Put-Cell $ws $kr 1 'Key: what these times are based on' $keyHead
    $keys = @(
        @('Tasks run one after another in the order they appear on the tracker. Nothing is assumed to run in parallel. If two tasks can run side by side, lower one of the Est. (hrs) values.', $null, $null),
        @('Hours per task come from Est. (hrs) on the tracker. Those come from the Sizing tab: the number of ESX hosts, clusters, vCenter Servers and NSX Edge nodes you enter there, multiplied by the hours per task type in the table. The table values are starting values, not measured timings. Change them to match your own experience.', $null, $null),
        @('Window length', ('"Each window is "&TEXT(Sizing!$F$7,"0.#")&" hours long. "&TEXT(Sizing!$F$5,"0%")&" of it is kept free at the end for overruns and validation, so "&TEXT(Sizing!$F$8,"0.0")&" hours of tasks are planned into each window."'), ('Each window is ' + $WIN_HRS + ' hours long. 20% of it is kept free at the end for overruns and validation, so ' + $WIN_USABLE.ToString('0.0', $script:Inv) + ' hours of tasks are planned into each window.')),
        @('A task is never split across two windows. If it does not fit in the hours left, it starts in the next window. A task longer than a whole window is counted in the Tasks longer than one window figure above and shows negative Spare hrs.', $null, $null),
        @('Window days', ('"Windows fall on "&IF(Sizing!$F$6="Yes","consecutive working days (Monday to Friday)","consecutive calendar days")&", starting from Window start on the tracker. The start time of day is taken from Window start."'), 'Windows fall on consecutive working days (Monday to Friday), starting from Window start on the tracker. The start time of day is taken from Window start.'),
        @('Not included: bank holidays and customer change freezes, approval and sign-off waits, data sync or backup waits outside the listed tasks, support calls, and rollback time. Check these against the customer calendar before proposing dates.', $null, $null),
        @('These are suggestions to open the conversation with the customer. They are not a commitment. After the first environment, compare Actual (hrs) with Est. (hrs) on the Dashboard and adjust the Sizing tab.', $null, $null)
    )
    $kk = $kr + 1
    foreach ($k in $keys) {
        $text = $k[0]
        if ($null -ne $k[1]) {
            Put-Cell $ws $kk 1 $k[1] $keyTxt 'f' $k[2] 's'
        } else {
            Put-Cell $ws $kk 1 $text $keyTxt
        }
        for ($c = 2; $c -le 10; $c++) { Put-Cell $ws $kk $c $null $keyTxt }
        [void]$ws.Merges.Add('A' + $kk + ':J' + $kk)
        $ws.RowH[$kk] = 32
        $kk++
    }
    Put-Cell $ws $kk 1 ('The table shows the first ' + $CW_ROWS + ' windows.') $keyI
    foreach ($cd in @(@(1,1,10), @(2,4,24), @(5,7,10), @(8,9,10), @(10,10,78))) { [void]$ws.Cols.Add($cd) }
    $cwWinCount = $winNo

    # ------------------------------------------------------------ Issues Log
    $ws = $iss
    Put-Cell $ws 1 1 'Issues Log' (Get-StyleId $st -sz 14 -b $true)
    Put-Cell $ws 2 1 ('One row per issue observed during the work. Ref points back to the ' + $tracker + ' item. The grey italic first row is an example; overwrite or delete it.') (Get-StyleId $st -i $true -color '595959')
    $ih = @('No.', 'Date / time', 'Ref', 'Issue observed', 'Severity', 'Action taken / workaround', 'Owner', 'Status', 'Resolved date / time', 'Vendor case / KB')
    for ($c = 1; $c -le $ih.Count; $c++) {
        Put-Cell $ws 4 $c $ih[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center' -wrap $true)
    }
    $ws.RowH[4] = 24
    $firstRef = '1.1'
    for ($pi = 1; $pi -le $phases.Count; $pi++) {
        if ($phases[$pi - 1].Name -ne 'Preparation') { $firstRef = [string]$pi + '.1'; break }
    }
    $serial = ($today.Date - (New-Object DateTime 1899, 12, 30)).Days
    $d1 = ($serial + 9.5 / 24).ToString('0.##########', $script:Inv)
    $d2 = ($serial + 10.25 / 24).ToString('0.##########', $script:Inv)
    $example = @{
        1  = @('1', 'n'); 2 = @($d1, 'n'); 3 = @($firstRef, 's')
        4  = @('Example: a precheck reports a stale DNS record for one host.', 's')
        5  = @('Amber', 's')
        6  = @('Corrected the record and re-ran the precheck. No impact on the window.', 's')
        7  = @('Engineer', 's'); 8 = @('Resolved', 's'); 9 = @($d2, 'n'); 10 = @('n/a', 's')
    }
    for ($r = 5; $r -le $ISS_LAST; $r++) {
        for ($c = 1; $c -le 10; $c++) {
            $ital = ($r -eq 5)
            $colr = $null; if ($r -eq 5) { $colr = '7F7F7F' }
            $nfx = 0; if ($c -in 2, 9) { $nfx = 165 }
            $hh = ''; if ($c -in 1, 2, 3, 5, 7, 8, 9) { $hh = 'center' }
            $sid = Get-StyleId $st -i $ital -color $colr -fill $INPUT -border $true -nf $nfx -h $hh -v 'center' -wrap $true
            if ($r -eq 5) { Put-Cell $ws $r $c $example[$c][0] $sid $example[$c][1] } else { Put-Cell $ws $r $c $null $sid }
        }
    }
    foreach ($cd in @(@(1,1,6), @(2,2,17), @(3,3,8), @(4,4,60), @(5,5,11), @(6,6,50), @(7,7,14), @(8,8,12), @(9,9,17), @(10,10,20))) { [void]$ws.Cols.Add($cd) }
    $ws.Freeze = @(0, 4)
    [void]$ws.Dv.Add(@(('E5:E' + $ISS_LAST), '"Green,Amber,Red"'))
    [void]$ws.Dv.Add(@(('H5:H' + $ISS_LAST), '"Open,Resolved"'))
    [void]$ws.Dv.Add(@(('C5:C' + $ISS_LAST), ($T + '!$A$' + $FIRST + ':$A$' + $LAST)))
    Add-Cf $ws ('E5:E' + $ISS_LAST) $rag

    # ------------------------------------------------------------ package
    $sheets = @($dash, $trk, $cw, $siz, $iss)
    $ns = $sheets.Count
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    $ct += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    $ct += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    for ($i = 1; $i -le $ns; $i++) { $ct += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' }
    $ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    $rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    $wbx += '<bookViews><workbookView activeTab="0"/></bookViews><sheets>'
    for ($i = 1; $i -le $ns; $i++) { $wbx += '<sheet name="' + (Esc-Xml $sheets[$i - 1].Name) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>' }
    $wbx += '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'
    $wrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for ($i = 1; $i -le $ns; $i++) { $wrels += '<Relationship Id="rId' + $i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $i + '.xml"/>' }
    $wrels += '<Relationship Id="rId' + ($ns + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add(@{ Name = '[Content_Types].xml';        Text = $ct })
    [void]$parts.Add(@{ Name = '_rels/.rels';                Text = $rels })
    [void]$parts.Add(@{ Name = 'xl/workbook.xml';            Text = $wbx })
    [void]$parts.Add(@{ Name = 'xl/_rels/workbook.xml.rels'; Text = $wrels })
    [void]$parts.Add(@{ Name = 'xl/styles.xml';              Text = (Get-StylesXml $st) })
    for ($i = 1; $i -le $ns; $i++) { [void]$parts.Add(@{ Name = ('xl/worksheets/sheet' + $i + '.xml'); Text = (Get-SheetXml $sheets[$i - 1]) }) }
    Write-Zip $outPath $parts
}

function Test-Xlsx([string]$path) {
    # Self-check: every part must be well-formed XML and every expected part must be present.
    $zip = Open-ZipRead $path
    try {
        $need = @('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
                  'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml', 'xl/worksheets/sheet5.xml')
        foreach ($n in $need) {
            $txt = Read-ZipText $zip $n
            if ($null -eq $txt) { throw ('Self-check failed: ' + $n + ' is missing from the workbook.') }
            $x = New-Object System.Xml.XmlDocument
            try { $x.LoadXml($txt) } catch { throw ('Self-check failed: ' + $n + ' is not well-formed XML. ' + $_.Exception.Message) }
        }
    } finally {
        $zip.Dispose()
    }
}

# ---------------------------------------------------------------------------
# Running it
# ---------------------------------------------------------------------------

function Convert-Plan([string]$docxPath, [string]$outPath) {
    $plan = Read-Plan $docxPath
    if (-not $outPath) {
        $outPath = Join-Path (Split-Path -Parent $docxPath) ([System.IO.Path]::GetFileNameWithoutExtension($docxPath) + ' - Tracker.xlsx')
    }
    $out = Get-UniquePath $outPath
    try {
        Build-Workbook $plan.Title $plan.Phases ([System.IO.Path]::GetFileName($docxPath)) $out $plan.Mode
        Test-Xlsx $out
    } catch {
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue }
        throw
    }
    return [pscustomobject]@{ Path = $out; Phases = $plan.Phases; Mode = $plan.Mode }
}

function Test-Ready($fi) {
    # True if the file looks finished being written and is not open in Word.
    $name = $fi.Name
    if ($name.StartsWith('~$')) { return $false }
    if ($name.Length -gt 2) {
        if (Test-Path -LiteralPath (Join-Path $fi.DirectoryName ('~$' + $name.Substring(2)))) { return $false }
    }
    if (Test-Path -LiteralPath (Join-Path $fi.DirectoryName ('~$' + $name))) { return $false }
    $fi.Refresh()
    $age = ((Get-Date) - $fi.LastWriteTime).TotalSeconds
    if ($age -lt 3) { return $false }
    $isZip = $false
    try {
        $fs = New-Object System.IO.FileStream($fi.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $buf = New-Object byte[] 4
            $n = $fs.Read($buf, 0, 4)
            if ($n -eq 4 -and $buf[0] -eq 0x50 -and $buf[1] -eq 0x4B -and $buf[2] -eq 3 -and $buf[3] -eq 4) { $isZip = $true }
        } finally { $fs.Dispose() }
    } catch {
        return $false
    }
    if ($isZip) { return $true }
    # Not a valid Word file. Give a copy that is still downloading a moment, then fail it properly.
    return ($age -gt 30)
}

function Process-Folder([string]$folder, [string]$outDir, [string]$doneDir, [string]$failedDir) {
    $count = 0
    $files = @(Get-ChildItem -LiteralPath $folder -Filter '*.docx' -File | Sort-Object Name)
    foreach ($fi in $files) {
        if (-not (Test-Ready $fi)) { continue }
        $stamp = (Get-Date).ToString('HH:mm:ss', $script:Inv)
        try {
            $res = Convert-Plan $fi.FullName (Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($fi.Name) + ' - Tracker.xlsx'))
            Move-Item -LiteralPath $fi.FullName -Destination (Get-UniquePath (Join-Path $doneDir $fi.Name))
            $items = 0; foreach ($p in $res.Phases) { $items += $p.Items.Count }
            Write-Host ($stamp + '  ' + $fi.Name + ' -> ' + (Split-Path -Leaf $res.Path) + '  (' + $res.Phases.Count + ' phases, ' + $items + ' items)' + $(if ($res.Mode -eq 'procedure') { '  [built from numbered steps]' } else { '' }))
            $count++
        } catch {
            $msg = $_.Exception.Message
            try {
                $dest = Get-UniquePath (Join-Path $failedDir $fi.Name)
                Move-Item -LiteralPath $fi.FullName -Destination $dest
                $note = [System.IO.Path]::ChangeExtension($dest, '.txt')
                [System.IO.File]::WriteAllText($note, ('Could not build a tracker from ' + $fi.Name + ':' + [Environment]::NewLine + $msg + [Environment]::NewLine), $script:Utf8)
            } catch {
                Write-Host ($stamp + '  ' + $fi.Name + ' could not be moved to Failed: ' + $_.Exception.Message)
            }
            Write-Host ($stamp + '  ' + $fi.Name + ' FAILED: ' + $msg)
        }
    }
    return $count
}

function Resolve-Full([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return [System.IO.Path]::GetFullPath($p) }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $p))
}

if ($Watch) {
    $folder = Resolve-Full $Watch
    $outDir = Join-Path $folder 'Trackers'
    $doneDir = Join-Path $folder 'Processed'
    $failedDir = Join-Path $folder 'Failed'
    foreach ($d in @($folder, $outDir, $doneDir, $failedDir)) {
        if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    }
    Write-Host ('Watching ' + $folder + '  (Ctrl+C to stop)')
    if ($Once) {
        [void](Process-Folder $folder $outDir $doneDir $failedDir)
    } else {
        while ($true) {
            [void](Process-Folder $folder $outDir $doneDir $failedDir)
            Start-Sleep -Seconds $Interval
        }
    }
} elseif ($Docx) {
    try {
        $src = Resolve-Full $Docx
        if (-not (Test-Path -LiteralPath $src)) { throw ('File not found: ' + $src) }
        $dest = ''
        if ($Output) { $dest = Resolve-Full $Output }
        $res = Convert-Plan $src $dest
        Write-Host ('Built ' + $res.Path)
        if ($res.Mode -eq 'procedure') { Write-Host 'No Execution Plan tables found, so this was built from the numbered steps under each main heading.' }
        $i = 0
        foreach ($ph in $res.Phases) { $i++; Write-Host ('  Phase ' + $i + ': ' + $ph.Name + ' (' + $ph.Items.Count + ' items)') }
    } catch {
        Write-Host ('FAILED: ' + $_.Exception.Message)
        exit 1
    }
} else {
    Write-Host 'Give a Word plan with -Docx "plan.docx", or a drop folder with -Watch "folder".'
    exit 1
}
