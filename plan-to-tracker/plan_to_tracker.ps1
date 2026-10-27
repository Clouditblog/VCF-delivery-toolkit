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
        if ($rows.Count -eq 0) { continue }
        $head = (Get-RowCells $rows[0] $nsm)
        $nonblank = @($head | Where-Object { $_ -ne '' })

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
    if (-not $hasExec) {
        throw 'No execution plan table found (a table headed Action / Impact to Environment / Required / VM Downtime).'
    }
    if (-not $title) { $title = [System.IO.Path]::GetFileNameWithoutExtension($path) }
    return [pscustomobject]@{ Title = $title; Phases = $phases }
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
    $x += '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy hh:mm"/></numFmts>'
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

function Build-Workbook([string]$title, $phases, [string]$sourceName, [string]$outPath) {
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
    $iss  = New-Sheet 'Issues Log'
    $dash.Grid = $false

    $NAVY = $script:NAVY; $INPUT = $script:INPUT; $STEP_A = $script:STEP_A; $STEP_B = $script:STEP_B

    # ------------------------------------------------------------ Tracker
    $ws = $trk
    Put-Cell $ws 1 1 ($title + ': Execution Tracker') (Get-StyleId $st -sz 14 -b $true)
    Put-Cell $ws 2 1 ('Source: ' + $sourceName + ', generated ' + $todayText + '. ' + $phases.Count + ' phases, ' + $itemsTotal + ' items. Items marked (optional) have Required = No in the plan.') (Get-StyleId $st -i $true -color '595959')
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
        @(4,  'Green', 'COUNTIF($G$[F]:$G$[L],"Green")', '00B050', 'FFFFFF', 0),
        @(5,  'Amber', 'COUNTIF($G$[F]:$G$[L],"Amber")', 'FFC000', $null, 0),
        @(6,  'Red',   'COUNTIF($G$[F]:$G$[L],"Red")', 'FF0000', 'FFFFFF', 0),
        @(7,  'Items complete (excl. step rows)', 'COUNTIF($F$[F]:$F$[L],"Complete")-COUNTIFS($B$[F]:$B$[L],"Step",$F$[F]:$F$[L],"Complete")', $null, $null, 0),
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
    Put-Cell $ws 10 1 'Yellow cells are for you to fill in. Pick Status and RAG from the dropdowns. Green = on plan, Amber = issue being managed, Red = blocked or needs a decision. Log each issue in the Issues Log tab using the item Ref.' (Get-StyleId $st -i $true -color '595959' -wrap $true -v 'top')
    $ws.RowH[10] = 54
    $heads = @('Ref', 'Type', 'Item', 'Owner', 'Planned date', 'Status', 'RAG', 'Actual start', 'Actual end', 'Notes / issues observed', 'Plan reference')
    for ($c = 1; $c -le $heads.Count; $c++) {
        Put-Cell $ws 11 $c $heads[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center' -wrap $true)
    }
    $ws.RowH[11] = 24
    $ctr = @(1, 2, 5, 6, 7, 8, 9)

    $r = $FIRST
    for ($pi = 1; $pi -le $phases.Count; $pi++) {
        $ph = $phases[$pi - 1]
        $vals = @{ 1 = ('S' + $pi); 2 = 'Step'; 3 = ('Phase ' + $pi + ': ' + $ph.Name); 6 = 'Not started' }
        for ($c = 1; $c -le 11; $c++) {
            $fillc = $STEP_B; if ($c -in 1, 2, 3, 11) { $fillc = $STEP_A }
            $hh = ''; if ($c -in 1, 2, 6) { $hh = 'center' }
            $sid = Get-StyleId $st -b $true -fill $fillc -border $true -h $hh -v 'center' -wrap $true
            Put-Cell $ws $r $c $vals[$c] $sid
        }
        $ws.RowH[$r] = 15
        $r++
        for ($ii = 1; $ii -le $ph.Items.Count; $ii++) {
            $it = $ph.Items[$ii - 1]
            $vals = @{ 1 = ([string]$pi + '.' + $ii); 2 = $it.Kind; 3 = $it.Text; 6 = 'Not started'; 11 = $it.Note }
            for ($c = 1; $c -le 11; $c++) {
                $nfx = 0; if ($c -eq 5) { $nfx = 164 } elseif ($c -in 8, 9) { $nfx = 165 }
                $fillc = $null; if ($c -in 4, 5, 6, 7, 8, 9, 10) { $fillc = $INPUT }
                $hh = ''; if ($ctr -contains $c) { $hh = 'center' }
                $sid = Get-StyleId $st -fill $fillc -border $true -nf $nfx -h $hh -v 'center' -wrap $true
                Put-Cell $ws $r $c $vals[$c] $sid
            }
            $ws.RowH[$r] = [math]::Max([double](Est-Height $it.Text 75), [double](Est-Height $it.Note 36))
            $r++
        }
    }
    for ($rr = $r; $rr -le $LAST; $rr++) {
        for ($c = 1; $c -le 11; $c++) {
            $nfx = 0; if ($c -eq 5) { $nfx = 164 } elseif ($c -in 8, 9) { $nfx = 165 }
            $hh = ''; if ($ctr -contains $c) { $hh = 'center' }
            Put-Cell $ws $rr $c $null (Get-StyleId $st -fill $INPUT -border $true -nf $nfx -h $hh -v 'center' -wrap $true)
        }
    }
    foreach ($cd in @(@(1,1,7), @(2,2,16), @(3,3,70), @(4,4,14), @(5,5,13), @(6,6,14), @(7,7,10), @(8,9,15), @(10,10,55), @(11,11,36))) { [void]$ws.Cols.Add($cd) }
    $ws.Freeze = @(3, 11)
    [void]$ws.Dv.Add(@(('F' + $FIRST + ':F' + $LAST), '"Not started,In progress,Complete,Blocked,Rolled back"'))
    [void]$ws.Dv.Add(@(('G' + $FIRST + ':G' + $LAST), '"Green,Amber,Red"'))
    Add-Cf $ws ('F' + $FIRST + ':F' + $LAST) @( @('equal', '"Complete"', $dComplete), @('equal', '"Blocked"', $dBlocked), @('equal', '"Rolled back"', $dBlocked) )
    Add-Cf $ws ('G' + $FIRST + ':G' + $LAST) $rag

    # ------------------------------------------------------------ Dashboard
    $ws = $dash
    Put-Cell $ws 1 1 ($title + ': Status at a glance') (Get-StyleId $st -sz 16 -b $true)
    $ws.RowH[1] = 20.65
    Put-Cell $ws 2 1 ('Updates itself from the ' + $tracker + ' and Issues Log tabs. Nothing to type on this page.') (Get-StyleId $st -i $true -color '595959')
    $tiles = @(
        @(1, 2,  'Overall complete', ($T + '!J9'),  9, 0),
        @(3, 4,  'Items done',       ($T + '!J7'),  0, 0),
        @(5, 6,  'Items total',      ($T + '!J8'),  0, $itemsTotal),
        @(7, 8,  'Open issues',      ($T + '!J10'), 0, 0),
        @(9, 10, 'Red items',        ($T + '!J6'),  0, 0)
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
    $dh = @('#', 'Phase', 'Status', 'RAG', 'Done', 'Total', '% done', 'Progress', 'Remaining')
    for ($c = 1; $c -le $dh.Count; $c++) {
        Put-Cell $ws 8 $c $dh[$c - 1] (Get-StyleId $st -b $true -color 'FFFFFF' -fill $NAVY -border $true -h 'center' -v 'center')
    }
    $ws.RowH[8] = 19.5
    $colA = $T + '!$A$' + $FIRST + ':$A$' + $LAST
    $colF = $T + '!$F$' + $FIRST + ':$F$' + $LAST
    $colG = $T + '!$G$' + $FIRST + ':$G$' + $LAST
    $colB = $T + '!$B$' + $FIRST + ':$B$' + $LAST
    $fStatus = 'IF(COUNTIFS([A],$A[r]&".*",[Fc],"Blocked")+COUNTIFS([A],$A[r]&".*",[Fc],"Rolled back")>0,"Blocked",IF(E[r]=F[r],"Complete",IF(OR(E[r]>0,COUNTIFS([A],$A[r]&".*",[Fc],"In progress")>0),"In progress","Not started")))'
    $fRag    = 'IF(OR(C[r]="Blocked",COUNTIFS([A],$A[r]&".*",[Gc],"Red")>0),"Red",IF(COUNTIFS([A],$A[r]&".*",[Gc],"Amber")>0,"Amber",IF(C[r]="Not started","","Green")))'
    $fDone   = 'COUNTIFS([A],$A[r]&".*",[Fc],"Complete")'
    $fTotal  = 'COUNTIFS([A],$A[r]&".*",[Bc],"<>Step")'
    $fBar    = 'REPT("' + $blk + '",ROUND(G[r]*20,0))&REPT("' + $lgt + '",20-ROUND(G[r]*20,0))'
    $barCached = $lgt * 20
    for ($i = 1; $i -le $phases.Count; $i++) {
        $ph = $phases[$i - 1]
        $r = 8 + $i
        $m = @{ '[A]' = $colA; '[Fc]' = $colF; '[Gc]' = $colG; '[Bc]' = $colB; '[r]' = $r }
        $ws.RowH[$r] = 19.5
        $nItems = $ph.Items.Count
        $base = Get-StyleId $st -border $true -h 'center' -v 'center'
        Put-Cell $ws $r 1 $i $base 'n'
        Put-Cell $ws $r 2 $ph.Name (Get-StyleId $st -border $true -h 'left' -v 'center')
        Put-Cell $ws $r 3 (Sub-Tok $fStatus $m) $base 'f' 'Not started' 's'
        Put-Cell $ws $r 4 (Sub-Tok $fRag $m) $base 'f' '' 's'
        Put-Cell $ws $r 5 (Sub-Tok $fDone $m) $base 'f' 0
        Put-Cell $ws $r 6 (Sub-Tok $fTotal $m) $base 'f' $nItems
        Put-Cell $ws $r 7 (Sub-Tok 'IF(F[r]=0,0,E[r]/F[r])' $m) (Get-StyleId $st -border $true -h 'center' -v 'center' -nf 9) 'f' 0
        Put-Cell $ws $r 8 (Sub-Tok $fBar $m) (Get-StyleId $st -color '2E75B6' -border $true -h 'left' -v 'center' -name 'Consolas') 'f' $barCached 's'
        Put-Cell $ws $r 9 (Sub-Tok 'F[r]-E[r]' $m) $base 'f' $nItems
    }
    $lastRow = 8 + $phases.Count
    Put-Cell $ws ($lastRow + 1) 2 'RAG per phase: Red if any item is Red or Blocked, Amber if any is Amber, otherwise Green once the phase has started.' (Get-StyleId $st -sz 9 -i $true -color '595959')
    Add-Cf $ws ('C9:C' + $lastRow) @( @('equal', '"Complete"', $dComplete), @('equal', '"In progress"', $dInProg), @('equal', '"Blocked"', $dBlocked), @('equal', '"Not started"', $dNotStarted) )
    Add-Cf $ws ('D9:D' + $lastRow) $rag
    foreach ($cd in @(@(1,1,5), @(2,2,42), @(3,3,14), @(4,4,10), @(5,6,8), @(7,7,9), @(8,8,26), @(9,9,11), @(10,10,3))) { [void]$ws.Cols.Add($cd) }

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
    $sheets = @($dash, $trk, $iss)
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    $ct += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    $ct += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    for ($i = 1; $i -le 3; $i++) { $ct += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' }
    $ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    $rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    $wbx += '<bookViews><workbookView activeTab="0"/></bookViews><sheets>'
    for ($i = 1; $i -le 3; $i++) { $wbx += '<sheet name="' + (Esc-Xml $sheets[$i - 1].Name) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>' }
    $wbx += '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'
    $wrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for ($i = 1; $i -le 3; $i++) { $wrels += '<Relationship Id="rId' + $i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $i + '.xml"/>' }
    $wrels += '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add(@{ Name = '[Content_Types].xml';        Text = $ct })
    [void]$parts.Add(@{ Name = '_rels/.rels';                Text = $rels })
    [void]$parts.Add(@{ Name = 'xl/workbook.xml';            Text = $wbx })
    [void]$parts.Add(@{ Name = 'xl/_rels/workbook.xml.rels'; Text = $wrels })
    [void]$parts.Add(@{ Name = 'xl/styles.xml';              Text = (Get-StylesXml $st) })
    for ($i = 1; $i -le 3; $i++) { [void]$parts.Add(@{ Name = ('xl/worksheets/sheet' + $i + '.xml'); Text = (Get-SheetXml $sheets[$i - 1]) }) }
    Write-Zip $outPath $parts
}

function Test-Xlsx([string]$path) {
    # Self-check: every part must be well-formed XML and every expected part must be present.
    $zip = Open-ZipRead $path
    try {
        $need = @('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
                  'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml')
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
        Build-Workbook $plan.Title $plan.Phases ([System.IO.Path]::GetFileName($docxPath)) $out
        Test-Xlsx $out
    } catch {
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue }
        throw
    }
    return [pscustomobject]@{ Path = $out; Phases = $plan.Phases }
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
            Write-Host ($stamp + '  ' + $fi.Name + ' -> ' + (Split-Path -Leaf $res.Path) + '  (' + $res.Phases.Count + ' phases, ' + $items + ' items)')
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
