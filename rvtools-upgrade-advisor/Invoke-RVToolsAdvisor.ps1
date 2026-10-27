<#
.SYNOPSIS
  Reads RVTools exports and produces an upgrade-readiness report (HTML + Excel) for VCF 9.1 or VVF 9.1.

.DESCRIPTION
  No modules needed. Works on Windows PowerShell 5.1 and PowerShell 7.
  Accepts one or more RVTools .xlsx files, or folders of them (or RVTools CSV exports).
  Each vCenter found becomes its own environment, with its own page in the HTML report.

.EXAMPLE
  .\Invoke-RVToolsAdvisor.ps1 -Path C:\rvtools\*.xlsx -Target 'VCF 9.1' -Features vSAN,NSX

.EXAMPLE
  .\Invoke-RVToolsAdvisor.ps1 -Path .\exports -Interactive

.EXAMPLE
  # Load a Broadcom Compatibility Guide server export so hosts are checked against it
  .\Invoke-RVToolsAdvisor.ps1 -ImportHcl .\bcg-servers.csv -HclTarget 'VCF 9.1'
#>
[CmdletBinding()]
param(
    [string[]]$Path,
    [ValidateSet('VCF 9.1', 'VVF 9.1')][string]$Target,
    [string[]]$Features,
    [string]$OutputPath = '.\RVTools-Upgrade-Report',
    [string]$Customer = '',
    [int]$SnapshotAgeDays = 7,
    [int]$DatastoreWarnPercent = 85,
    [switch]$Interactive,
    [string]$ImportHcl,
    [string]$HclTarget
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$script:Root = $PSScriptRoot
if (-not $script:Root) { $script:Root = (Get-Location).Path }
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
# allow -Features vSAN,NSX and -Path a.xlsx,b.xlsx even when passed through -File
$Features = @($Features | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$Path = @($Path | ForEach-Object { $_ -split ',(?=\s*[A-Za-z]:|\s*\.|\s*/|\s*\\)' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

#region ---------- helpers ----------
function Norm([string]$s) { if ($null -eq $s) { return '' }; ($s.ToLowerInvariant() -replace '[^a-z0-9]', '') }

function V($row, [string[]]$names) {
    foreach ($n in $names) {
        if ($row.ContainsKey($n)) { $v = $row[$n]; if ($null -ne $v -and "$v" -ne '') { return $v } }
    }
    return $null
}
function N($row, [string[]]$names) {
    $v = V $row $names
    if ($null -eq $v) { return 0.0 }
    $d = 0.0
    if ([double]::TryParse("$v", [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $d }
    return 0.0
}
function IsTrue($v) { return ("$v" -match '^(true|yes|1|enabled)$') }
function ToDate($v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    $d = 0.0
    if ([double]::TryParse("$v", [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$d) -and $d -gt 20000 -and $d -lt 80000) { return [DateTime]::FromOADate($d) }
    $dt = [DateTime]::MinValue
    if ([DateTime]::TryParse("$v", [ref]$dt)) { return $dt }
    return $null
}
function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode("$s") }
function GiB($mib) { [math]::Round($mib / 1024, 1) }

function Parse-EsxVersion([string]$s) {
    $o = [ordered]@{ Text = $s; Major = 0; Minor = 0; Patch = 0; Label = $s; Known = $false }
    if ($s -match '(\d+)\.(\d+)(?:\.(\d+))?') {
        $o.Major = [int]$Matches[1]; $o.Minor = [int]$Matches[2]
        if ($Matches[3]) { $o.Patch = [int]$Matches[3] }
        $o.Known = $true
        $upd = if ($o.Patch -gt 0 -and $o.Major -ge 7) { " U$($o.Patch)" } else { '' }
        $o.Label = "$($o.Major).$($o.Minor)$upd"
    }
    [pscustomobject]$o
}
#endregion

#region ---------- xlsx reader (no modules) ----------
function Read-XlsxWorkbook([string]$file) {
    $fs = [System.IO.File]::Open($file, 'Open', 'Read', 'ReadWrite')
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $getEntry = { param($n) $zip.Entries | Where-Object { $_.FullName -eq $n } | Select-Object -First 1 }
        # shared strings
        $shared = New-Object System.Collections.Generic.List[string]
        $e = & $getEntry 'xl/sharedStrings.xml'
        if ($e) {
            $st = $e.Open(); $r = [System.Xml.XmlReader]::Create($st)
            $sb = $null
            while ($r.Read()) {
                if ($r.NodeType -eq 'Element') {
                    if ($r.LocalName -eq 'si') { $sb = New-Object System.Text.StringBuilder; if ($r.IsEmptyElement) { $shared.Add(''); $sb = $null } }
                    elseif ($r.LocalName -eq 't' -and $sb -ne $null) { [void]$sb.Append($r.ReadElementContentAsString()); }
                }
                if ($r.NodeType -eq 'EndElement' -and $r.LocalName -eq 'si' -and $sb -ne $null) { $shared.Add($sb.ToString()); $sb = $null }
            }
            $r.Close(); $st.Close()
        }
        # sheet names -> paths
        $wb = New-Object System.Xml.XmlDocument
        $s = (& $getEntry 'xl/workbook.xml').Open(); $wb.Load($s); $s.Close()
        $rel = New-Object System.Xml.XmlDocument
        $s = (& $getEntry 'xl/_rels/workbook.xml.rels').Open(); $rel.Load($s); $s.Close()
        $relMap = @{}
        foreach ($n in $rel.DocumentElement.ChildNodes) { if ($n.Id) { $relMap[$n.Id] = $n.Target } }
        $result = @{}
        foreach ($sh in $wb.GetElementsByTagName('sheet')) {
            $name = $sh.GetAttribute('name')
            $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = $relMap[$rid]
            if (-not $target) { continue }
            $target = $target.TrimStart('/'); if ($target -notmatch '^xl/') { $target = "xl/$target" }
            $entry = & $getEntry $target
            if (-not $entry) { continue }
            $result[$name] = Read-SheetXml $entry $shared
        }
        return $result
    }
    finally { $zip.Dispose(); $fs.Dispose() }
}

function ColIndex([string]$ref) {
    $n = 0
    foreach ($ch in $ref.ToCharArray()) { if ($ch -ge 'A' -and $ch -le 'Z') { $n = $n * 26 + ([int]$ch - 64) } else { break } }
    return $n - 1
}

function Read-SheetXml($entry, $shared) {
    $rows = New-Object System.Collections.Generic.List[object]
    $st = $entry.Open(); $r = [System.Xml.XmlReader]::Create($st)
    $cells = $null; $col = 0; $type = ''; $inIs = $false; $isText = $null
    $more = $r.Read()
    while ($more) {
        $advance = $true
        if ($r.NodeType -eq 'Element') {
            switch ($r.LocalName) {
                'row' { $cells = @{} }
                'c' {
                    $col = ColIndex $r.GetAttribute('r'); $type = $r.GetAttribute('t')
                }
                'v' {
                    $txt = $r.ReadElementContentAsString(); $advance = $false
                    if ($type -eq 's') { $txt = $shared[[int]$txt] }
                    elseif ($type -eq 'b') { $txt = if ($txt -eq '1') { 'True' } else { 'False' } }
                    if ($cells -ne $null) { $cells[$col] = $txt }
                }
                'is' { $inIs = $true; $isText = New-Object System.Text.StringBuilder }
                't' { if ($inIs) { [void]$isText.Append($r.ReadElementContentAsString()); $advance = $false } }
            }
        }
        elseif ($r.NodeType -eq 'EndElement') {
            if ($r.LocalName -eq 'is') { $inIs = $false; if ($cells -ne $null) { $cells[$col] = $isText.ToString() } }
            elseif ($r.LocalName -eq 'row') { if ($cells -ne $null -and $cells.Count -gt 0) { $rows.Add($cells) }; $cells = $null }
        }
        if ($advance) { $more = $r.Read() } else { $more = -not $r.EOF }
    }
    $r.Close(); $st.Close()
    if ($rows.Count -lt 1) { return @() }
    $hdr = @{}
    foreach ($k in $rows[0].Keys) { $hdr[$k] = Norm ([string]$rows[0][$k]) }
    $out = New-Object System.Collections.Generic.List[hashtable]
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $h = @{}
        foreach ($k in $hdr.Keys) { if ($hdr[$k] -and $rows[$i].ContainsKey($k)) { $h[$hdr[$k]] = $rows[$i][$k] } }
        if ($h.Count -gt 0) { $out.Add($h) }
    }
    return , $out
}

function Read-CsvSet([string[]]$files) {
    $result = @{}
    foreach ($f in $files) {
        if ($f -match 'tab(\w+)\.csv$') { $name = $Matches[1] } else { continue }
        $list = New-Object System.Collections.Generic.List[hashtable]
        foreach ($row in (Import-Csv -LiteralPath $f)) {
            $h = @{}
            foreach ($p in $row.PSObject.Properties) { $h[(Norm $p.Name)] = $p.Value }
            $list.Add($h)
        }
        $result[$name] = $list
    }
    return $result
}

function Get-Sheet($wb, [string]$name) {
    foreach ($k in $wb.Keys) { if ($k -ieq $name) { return $wb[$k] } }
    return @()
}
#endregion

#region ---------- xlsx writer (no modules) ----------
function XmlEsc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}
function ColName([int]$i) { $s = ''; $i++; while ($i -gt 0) { $m = ($i - 1) % 26; $s = [char](65 + $m) + $s; $i = [math]::Floor(($i - 1) / 26) }; $s }
function SafeSheetName([string]$n, $used) {
    $n = ($n -replace '[\\/\?\*\[\]:]', '-'); if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
    $base = $n; $i = 2
    while ($used -contains $n.ToLower()) { $suffix = "~$i"; $n = $base.Substring(0, [math]::Min($base.Length, 31 - $suffix.Length)) + $suffix; $i++ }
    return $n
}

# sheets: array of @{ Name; Headers; Rows (list of object[]); Widths; SevCols; WrapCols; Links (hashtable "r,c" -> target sheet) }
function Write-Xlsx([string]$file, $sheets) {
    if (Test-Path $file) { Remove-Item $file -Force }
    $fs = [System.IO.File]::Open($file, 'Create')
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $add = {
        param($name, $content)
        $e = $zip.CreateEntry($name)
        $w = New-Object System.IO.StreamWriter($e.Open(), $utf8)
        $w.Write($content); $w.Dispose()
    }
    try {
        $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        for ($i = 1; $i -le $sheets.Count; $i++) { $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" }
        $ct += '</Types>'
        & $add '[Content_Types].xml' $ct
        & $add '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        $wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        for ($i = 1; $i -le $sheets.Count; $i++) {
            $wbx += "<sheet name=`"$(XmlEsc $sheets[$i-1].Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
            $rels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
        }
        $wbx += '</sheets></workbook>'
        $rels += "<Relationship Id=`"rId$($sheets.Count + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>"
        & $add 'xl/workbook.xml' $wbx
        & $add 'xl/_rels/workbook.xml.rels' $rels
        $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/></font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>
'@
        & $add 'xl/styles.xml' $styles
        for ($i = 1; $i -le $sheets.Count; $i++) {
            $sh = $sheets[$i - 1]
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
            $hasHdr = ($sh.Headers -and $sh.Headers.Count -gt 0)
            $hdrRow = if ($sh.HeaderRow) { [int]$sh.HeaderRow } else { 1 }
            if ($hasHdr -and $sh.Freeze -ne $false) { [void]$sb.Append("<sheetViews><sheetView workbookViewId=`"0`"><pane ySplit=`"$hdrRow`" topLeftCell=`"A$($hdrRow + 1)`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews>") }
            if ($sh.Widths) {
                [void]$sb.Append('<cols>')
                for ($c = 0; $c -lt $sh.Widths.Count; $c++) { [void]$sb.Append("<col min=`"$($c + 1)`" max=`"$($c + 1)`" width=`"$($sh.Widths[$c])`" customWidth=`"1`"/>") }
                [void]$sb.Append('</cols>')
            }
            [void]$sb.Append('<sheetData>')
            $rowNo = 0
            foreach ($pre in $sh.PreRows) {   # title lines above the table
                if ($null -eq $pre) { continue }
                $rowNo++
                $sty = if ($pre.Style) { $pre.Style } else { 0 }
                [void]$sb.Append("<row r=`"$rowNo`"><c r=`"A$rowNo`" s=`"$sty`" t=`"inlineStr`"><is><t>$(XmlEsc $pre.Text)</t></is></c></row>")
            }
            if ($hasHdr) {
                $rowNo++
                [void]$sb.Append("<row r=`"$rowNo`" ht=`"22`" customHeight=`"1`">")
                for ($c = 0; $c -lt $sh.Headers.Count; $c++) { [void]$sb.Append("<c r=`"$(ColName $c)$rowNo`" s=`"1`" t=`"inlineStr`"><is><t>$(XmlEsc $sh.Headers[$c])</t></is></c>") }
                [void]$sb.Append('</row>')
            }
            $first = $rowNo
            foreach ($row in $sh.Rows) {
                $rowNo++
                [void]$sb.Append("<row r=`"$rowNo`">")
                for ($c = 0; $c -lt $row.Count; $c++) {
                    $v = $row[$c]
                    $ref = "$(ColName $c)$rowNo"
                    if ($null -eq $v -or "$v" -eq '') { continue }
                    $sty = 0
                    $isNum = ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal])
                    if (-not $isNum) {
                        if ($sh.SevCols -and ($sh.SevCols -contains $c)) {
                            switch -regex ("$v") { '^(Blocker|Blocked)$' { $sty = 2 } '^(Warning|Ready with actions)$' { $sty = 3 } '^(OK|Ready)$' { $sty = 4 } '^Info$' { $sty = 5 } }
                        }
                        elseif ($sh.WrapCols -and ($sh.WrapCols -contains $c)) { $sty = 6 }
                        if ($sh.LinkCols -and ($sh.LinkCols -contains $c)) { $sty = 7 }
                    }
                    if ($isNum) { [void]$sb.Append("<c r=`"$ref`" s=`"$sty`"><v>$([string]::Format($script:Inv, '{0}', $v))</v></c>") }
                    else { [void]$sb.Append("<c r=`"$ref`" s=`"$sty`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(XmlEsc "$v")</t></is></c>") }
                }
                [void]$sb.Append('</row>')
            }
            [void]$sb.Append('</sheetData>')
            if ($hasHdr -and $sh.Rows.Count -gt 0 -and -not $sh.NoFilter) {
                [void]$sb.Append("<autoFilter ref=`"A${hdrRow}:$(ColName ($sh.Headers.Count - 1))$rowNo`"/>")
            }
            if ($sh.LinkTargets -and $sh.LinkTargets.Count -gt 0) {
                [void]$sb.Append('<hyperlinks>')
                foreach ($k in $sh.LinkTargets.Keys) { [void]$sb.Append("<hyperlink ref=`"$k`" location=`"'$(XmlEsc $sh.LinkTargets[$k])'!A1`" display=`"$(XmlEsc $sh.LinkTargets[$k])`"/>") }
                [void]$sb.Append('</hyperlinks>')
            }
            [void]$sb.Append('</worksheet>')
            & $add "xl/worksheets/sheet$i.xml" $sb.ToString()
        }
    }
    finally { $zip.Dispose(); $fs.Dispose() }
}
#endregion

#region ---------- config, HCL ----------
function Load-Config {
    $f = Join-Path $script:Root 'data\targets.json'
    if (-not (Test-Path $f)) { $f = Join-Path $script:Root 'data/targets.json' }
    if (-not (Test-Path $f)) { throw "Cannot find data\targets.json next to the script." }
    $script:Cfg = Get-Content -LiteralPath $f -Raw | ConvertFrom-Json
}
function Get-HclFile { Join-Path $script:Root 'data/hcl-servers.csv' }

function NormModel([string]$s) { ((Norm $s) -replace 'dell(emc|inc)?|hpe?|hewlettpackard(enterprise)?|lenovo|cisco|fujitsu|supermicro|inspur|huawei|poweredge', '') }

function Import-HclExport([string]$csv, [string]$target) {
    if (-not $target) { throw 'Use -HclTarget with -ImportHcl (for example -HclTarget "VCF 9.1").' }
    $ver = if ($target -match '(\d+\.\d+)') { $Matches[1] } else { $target }
    $rows = Import-Csv -LiteralPath $csv
    if (-not $rows) { throw "No rows in $csv" }
    $props = $rows[0].PSObject.Properties.Name
    $mc = $props | Where-Object { $_ -match 'model|system|server|product' } | Select-Object -First 1
    $vc = $props | Where-Object { $_ -match 'vendor|partner|brand|manufacturer' } | Select-Object -First 1
    $rc = $props | Where-Object { $_ -match 'release|version|supported' } | Select-Object -First 1
    if (-not $mc) { throw "Could not find a model column in $csv. Columns: $($props -join ', ')" }
    $out = New-Object System.Collections.Generic.List[object]
    $existing = @(); if (Test-Path (Get-HclFile)) { $existing = @(Import-Csv (Get-HclFile)) }
    $seen = @{}; foreach ($e in $existing) { $seen["$($e.Vendor)|$($e.Model)|$($e.Target)"] = $true }
    $added = 0
    foreach ($r in $rows) {
        if ($rc -and ("$($r.$rc)" -notmatch [regex]::Escape($ver))) { continue }
        $vendor = if ($vc) { "$($r.$vc)" } else { '' }
        $key = "$vendor|$($r.$mc)|$target"
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $out.Add([pscustomobject]@{ Vendor = $vendor; Model = "$($r.$mc)"; Target = $target; Status = 'Supported'; Source = (Split-Path $csv -Leaf); Imported = (Get-Date -Format 'yyyy-MM-dd') })
        $added++
    }
    (@($existing) + $out.ToArray()) | Export-Csv -LiteralPath (Get-HclFile) -NoTypeInformation
    Write-Host "Imported $added supported server model rows for $target into data\hcl-servers.csv"
}

function Get-HclModels([string]$target) {
    $f = Get-HclFile
    if (-not (Test-Path $f)) { return @() }
    @(Import-Csv $f | Where-Object { $_.Target -eq $target } | ForEach-Object {
            [pscustomobject]@{ Norm = (NormModel $_.Model); Vendor = $_.Vendor; Model = $_.Model }
        })
}
function Find-Hcl($hcl, [string]$vendor, [string]$model) {
    $m = NormModel $model
    if (-not $m) { return $null }
    foreach ($h in $hcl) { if ($h.Norm -and ($m -eq $h.Norm -or $m.Contains($h.Norm) -or $h.Norm.Contains($m))) { return $h } }
    return $null
}
#endregion

#region ---------- analysis ----------
function Add-Finding($list, [string]$envName, [string]$sev, [string]$area, [string]$obj, [string]$finding, [string]$rec) {
    $list.Add([pscustomobject]@{ Environment = $envName; Severity = $sev; Area = $area; Object = $obj; Finding = $finding; Recommendation = $rec })
}
function Bump-Sev([string]$cur, [string]$new) {
    $rank = @{ 'OK' = 0; 'Info' = 1; 'Warning' = 2; 'Blocker' = 3 }
    if ($rank[$new] -gt $rank[$cur]) { return $new } else { return $cur }
}
function Sample($names, [int]$n = 5) {
    $a = @($names); $s = ($a | Select-Object -First $n) -join ', '
    if ($a.Count -gt $n) { $s += " and $($a.Count - $n) more" }
    $s
}

function Analyze-Environment([string]$envName, $d, $t, [string]$targetName, [string[]]$features, $hcl) {
    $F = New-Object System.Collections.Generic.List[object]
    $minSrc = [version]$t.minSourceEsxi
    $ruleSet = @($script:Cfg.cpuRuleSets.($t.cpuRules))

    # ---- vCenter
    $vc = @()
    foreach ($s in $d.vSource) {
        $full = "$(V $s @('fullname','version'))"; $ver = Parse-EsxVersion "$(V $s @('version','apiversion'))"
        $vc += [pscustomobject]@{ Name = "$(V $s @('name','visdkserver'))"; Version = $ver; Full = $full; Build = "$(V $s @('build'))"; OsType = "$(V $s @('ostype'))" }
    }
    if ($vc.Count -eq 0) {
        Add-Finding $F $envName 'Info' 'vCenter' $envName 'No vSource tab, so the vCenter version is unknown.' 'Re-export from RVTools with all tabs, or confirm the vCenter version with the customer.'
    }
    foreach ($v in $vc) {
        if ($v.Version.Known -and $v.Version.Major -lt $minSrc.Major) {
            Add-Finding $F $envName 'Blocker' 'vCenter' $v.Name "vCenter $($v.Version.Label) is below the supported starting point ($($t.minSourceEsxi))." 'Upgrade to vSphere 8.x first. There is no direct path to 9.x from here.'
        }
    }

    # ---- hosts
    $hosts = New-Object System.Collections.Generic.List[object]
    foreach ($h in $d.vHost) {
        $name = "$(V $h @('vmhost','host','name'))"
        $cpm = "$(V $h @('cpumodel'))"
        $sock = [int](N $h @('cpu', 'cpus', 'numcpu')); $cpp = [int](N $h @('corespercpu'))
        $cores = [int](N $h @('cores', 'numcores')); if ($cores -eq 0) { $cores = $sock * $cpp }
        if ($cpp -eq 0 -and $sock -gt 0) { $cpp = [int]($cores / $sock) }
        $memMiB = N $h @('memory', 'nummemory')
        $esx = Parse-EsxVersion "$(V $h @('esxversion'))"
        $vendor = "$(V $h @('vendor'))"; $model = "$(V $h @('model'))"
        $cls = $null
        foreach ($r in $ruleSet) { if ($cpm -match $r.pattern) { $cls = $r; break } }
        $cpuStatus = 'Verify'; $cpuFam = 'Not in bundled rules'
        if ($cls) { $cpuStatus = $cls.status; $cpuFam = $cls.family }
        elseif ($ruleSet.Count -eq 0) { $cpuFam = 'No CPU rules for this target' }
        $lic = $sock * [math]::Max(16, $cpp)
        $hosts.Add([pscustomobject]@{
                Environment = $envName; Host = $name; Cluster = "$(V $h @('cluster'))"; Vendor = $vendor; Model = $model; CpuModel = $cpm
                Sockets = $sock; CoresPerCpu = $cpp; Cores = $cores; LicensedCores = $lic; MemGiB = (GiB $memMiB); VMs = [int](N $h @('vms', 'numvms'))
                EsxVersion = $esx.Label; EsxMajor = $esx.Major; EsxKnown = $esx.Known; CpuStatus = $cpuStatus; CpuFamily = $cpuFam; Hcl = ''; Severity = 'OK'; Notes = ''
            })
    }
    foreach ($h in $hosts) {
        $notes = New-Object System.Collections.Generic.List[string]; $sev = 'OK'
        if ($h.EsxKnown -and $h.EsxMajor -lt $minSrc.Major) {
            Add-Finding $F $envName 'Blocker' 'Host' $h.Host "ESXi $($h.EsxVersion) cannot be upgraded straight to $targetName." "Upgrade to ESXi $($t.minSourceEsxi) or later first. vCenter 9 cannot manage ESX 7.0 or earlier."
            $notes.Add("ESXi $($h.EsxVersion) needs an interim upgrade"); $sev = Bump-Sev $sev 'Blocker'
        }
        elseif ($h.EsxKnown -and $h.EsxMajor -ge $t.esxiMajor) { $notes.Add("Already on ESXi $($h.EsxVersion)") }
        switch ($h.CpuStatus) {
            'Blocker' { Add-Finding $F $envName 'Blocker' 'Host' $h.Host "CPU '$($h.CpuModel)' ($($h.CpuFamily)) is not supported by ESX 9.x." 'Replace or retire this host before the upgrade. Plan hardware refresh into the project.'; $notes.Add("CPU blocked: $($h.CpuFamily)"); $sev = Bump-Sev $sev 'Blocker' }
            'Warning' { $why = ($ruleSet | Where-Object { $_.family -eq $h.CpuFamily } | Select-Object -First 1).note; Add-Finding $F $envName 'Warning' 'Host' $h.Host "CPU '$($h.CpuModel)' is $($h.CpuFamily). $why" 'Check the exact server model on the Broadcom Compatibility Guide and plan a refresh horizon.'; $notes.Add("CPU: $($h.CpuFamily)"); $sev = Bump-Sev $sev 'Warning' }
            'Verify' { if ($ruleSet.Count -gt 0) { Add-Finding $F $envName 'Info' 'Host' $h.Host "CPU '$($h.CpuModel)' is not covered by the bundled rules." 'Check the CPU on the Broadcom Compatibility Guide.'; $sev = Bump-Sev $sev 'Info' } }
        }
        # HCL
        if ($hcl.Count -gt 0) {
            $m = Find-Hcl $hcl $h.Vendor $h.Model
            if ($m) { $h.Hcl = 'Listed' } else {
                $h.Hcl = 'Not listed'
                Add-Finding $F $envName 'Warning' 'Host' $h.Host "Server model '$($h.Vendor) $($h.Model)' was not found in the loaded HCL data for $targetName." 'Search the model on the Broadcom Compatibility Guide. The HCL data may be incomplete or the model name may differ.'
                $notes.Add('Model not in loaded HCL'); $sev = Bump-Sev $sev 'Warning'
            }
        }
        else { $h.Hcl = 'Not checked' }
        if ($h.Sockets -gt 0 -and $h.CoresPerCpu -lt 16) { $notes.Add("Licensed at 16 cores per CPU (has $($h.CoresPerCpu))") }
        $h.Severity = $sev; $h.Notes = ($notes -join '; ')
    }

    # ---- clusters
    $clusters = New-Object System.Collections.Generic.List[object]
    $cInfo = @{}; foreach ($c in $d.vCluster) { $cInfo["$(V $c @('name'))"] = $c }
    foreach ($g in ($hosts | Group-Object Cluster)) {
        $cn = if ($g.Name) { $g.Name } else { '(standalone)' }
        $ci = $cInfo[$g.Name]
        $ha = ''; $drs = ''; $vsan = ''
        if ($ci) { $ha = if (IsTrue (V $ci @('haenabled'))) { 'On' } else { 'Off' }; $drs = if (IsTrue (V $ci @('drsenabled'))) { 'On' } else { 'Off' }; $vsan = if (IsTrue (V $ci @('vsanenabled', 'vsan'))) { 'On' } else { 'Off' } }
        $vers = @($g.Group | Select-Object -ExpandProperty EsxVersion -Unique)
        $blk = @($g.Group | Where-Object { $_.Severity -eq 'Blocker' }).Count
        $sev = if ($blk -gt 0) { 'Blocker' } elseif (@($g.Group | Where-Object { $_.Severity -eq 'Warning' }).Count -gt 0) { 'Warning' } else { 'OK' }
        $clusters.Add([pscustomobject]@{ Environment = $envName; Cluster = $cn; Hosts = $g.Count; VMs = ($g.Group | Measure-Object VMs -Sum).Sum; Cores = ($g.Group | Measure-Object Cores -Sum).Sum; MemGiB = [math]::Round(($g.Group | Measure-Object MemGiB -Sum).Sum, 0); HA = $ha; DRS = $drs; vSAN = $vsan; EsxVersions = ($vers -join ', '); BlockedHosts = $blk; Severity = $sev })
        if ($g.Count -eq 1 -and $g.Name) { Add-Finding $F $envName 'Warning' 'Cluster' $cn 'Single-host cluster. There is nowhere to move VMs while the host is upgraded.' 'Plan VM downtime or add temporary capacity.' }
        if ($vers.Count -gt 1) { Add-Finding $F $envName 'Info' 'Cluster' $cn "Mixed ESXi versions in the cluster: $($vers -join ', ')." 'Bring the cluster to one 8.x version before starting the 9.1 upgrade.' }
        if ($drs -eq 'Off' -and $g.Count -gt 1) { Add-Finding $F $envName 'Info' 'Cluster' $cn 'DRS is off. Rolling host upgrades will need manual VM evacuation.' 'Turn DRS on for the upgrade window, or plan manual vMotion.' }
    }

    # ---- datastores
    $stores = New-Object System.Collections.Generic.List[object]
    foreach ($s in $d.vDatastore) {
        $cap = N $s @('capacitymib'); $free = N $s @('freemib'); $used = if ($cap -gt 0) { [math]::Round(($cap - $free) * 100 / $cap, 0) } else { 0 }
        $vmfs = "$(V $s @('vmfsversion','version'))"
        $stores.Add([pscustomobject]@{ Environment = $envName; Datastore = "$(V $s @('name'))"; Type = "$(V $s @('type'))"; CapacityGiB = [math]::Round($cap / 1024, 0); FreeGiB = [math]::Round($free / 1024, 0); UsedPct = $used; VmfsVersion = $vmfs; Severity = 'OK' })
    }
    foreach ($s in $stores) {
        if ($s.UsedPct -ge $DatastoreWarnPercent) { $s.Severity = 'Warning' }
    }
    $full = @($stores | Where-Object { $_.Severity -eq 'Warning' })
    if ($full.Count -gt 0) { Add-Finding $F $envName 'Warning' 'Datastore' (Sample $full.Datastore) "$($full.Count) datastore(s) are at or above $DatastoreWarnPercent% used." 'Free space before the upgrade. Snapshots and vMotion during patching need headroom.' }
    $oldVmfs = @($stores | Where-Object { $_.VmfsVersion -match '^[1-5]\b' -and $_.Type -match 'VMFS' })
    if ($oldVmfs.Count -gt 0) { Add-Finding $F $envName 'Info' 'Datastore' (Sample $oldVmfs.Datastore) "$($oldVmfs.Count) datastore(s) are on VMFS 5 or older." 'Plan a move to VMFS 6 (new datastore and Storage vMotion).' }

    # ---- VMs
    $snapBy = @{}; foreach ($s in $d.vSnapshot) { $k = "$(V $s @('vm'))"; if (-not $snapBy.ContainsKey($k)) { $snapBy[$k] = New-Object System.Collections.Generic.List[object] }; $snapBy[$k].Add($s) }
    $toolBy = @{}; foreach ($s in $d.vTools) { $toolBy["$(V $s @('vm'))"] = $s }
    $rdmBy = @{}; foreach ($s in $d.vDisk) { if (IsTrue (V $s @('raw'))) { $rdmBy["$(V $s @('vm'))"] = $true } }
    $vms = New-Object System.Collections.Generic.List[object]
    $eol = $script:Cfg.guestOsEndOfSupport
    $hwMin = [int]$t.minHwAdvisory
    $now = Get-Date
    foreach ($v in $d.vInfo) {
        $vn = "$(V $v @('vm'))"
        $power = "$(V $v @('powerstate'))"; $tpl = IsTrue (V $v @('template'))
        $hwRaw = "$(V $v @('hwversion'))"; $hw = 0; if ($hwRaw -match '(\d+)') { $hw = [int]$Matches[1] }
        $os = "$(V $v @('osaccordingtotheconfigurationfile','osaccordingtothevmwaretools'))"
        $memMiB = N $v @('memory')
        $flags = New-Object System.Collections.Generic.List[string]; $sev = 'OK'
        $snapCount = 0; $oldest = $null
        if ($snapBy.ContainsKey($vn)) {
            $snapCount = $snapBy[$vn].Count
            foreach ($sn in $snapBy[$vn]) { $dt = ToDate (V $sn @('datetime')); if ($dt -and (-not $oldest -or $dt -lt $oldest)) { $oldest = $dt } }
            $age = if ($oldest) { [int]($now - $oldest).TotalDays } else { -1 }
            if ($age -ge $SnapshotAgeDays -or $age -lt 0) { $flags.Add("Old snapshot$(if ($age -ge 0) { " ($age days)" } else { ' (age unknown)' })"); $sev = 'Warning' }
        }
        if ($hw -gt 0 -and $hw -lt $hwMin) { $flags.Add("Hardware version $hw (below $hwMin)"); $sev = 'Warning' }
        if ($rdmBy.ContainsKey($vn)) { $flags.Add('Raw device mapping'); $sev = 'Warning' }
        if ($eol -and $os -match $eol) { $flags.Add('Guest OS past vendor end of support'); $sev = 'Warning' }
        if (IsTrue (V $v @('consolidationneeded'))) { $flags.Add('Snapshot consolidation needed'); $sev = 'Warning' }
        $toolTxt = ''
        if ($toolBy.ContainsKey($vn)) {
            $toolTxt = "$(V $toolBy[$vn] @('tools','toolsstatus'))"
            if ($power -match 'On' -and -not $tpl) {
                if ($toolTxt -match 'NotInstalled') { $flags.Add('VMware Tools not installed'); $sev = 'Warning' }
                elseif ($toolTxt -match 'NeedUpgrade|Old|TooOld|Blacklisted' -or (IsTrue (V $toolBy[$vn] @('upgradeable')))) { $flags.Add('VMware Tools outdated'); if ($sev -eq 'OK') { $sev = 'Info' } }
            }
        }
        $vms.Add([pscustomobject]@{ Environment = $envName; VM = $vn; Power = $power; Template = $tpl; Cluster = "$(V $v @('cluster'))"; Host = "$(V $v @('host'))"; vCPU = [int](N $v @('cpus')); MemGiB = (GiB $memMiB); HwVersion = $hw; Tools = $toolTxt; GuestOS = $os; Snapshots = $snapCount; Severity = $sev; Flags = ($flags -join '; ') })
    }
    $vmList = @($vms | Where-Object { -not $_.Template })
    $vmRules = @(
        @{ Key = 'Old snapshot'; Sev = 'Warning'; Text = 'VM(s) have snapshots older than the threshold or of unknown age.'; Rec = 'Delete or consolidate snapshots before the upgrade. Old snapshots slow vMotion and break backups.' },
        @{ Key = 'Hardware version'; Sev = 'Warning'; Text = "VM(s) run a very old virtual hardware version (below $hwMin)."; Rec = 'Upgrade VMware Tools first, then raise hardware version in a change window. Check the guest OS is still supported.' },
        @{ Key = 'Raw device'; Sev = 'Warning'; Text = 'VM(s) use raw device mappings.'; Rec = 'List RDMs early. They can block or slow host evacuation and need storage-team review.' },
        @{ Key = 'end of support'; Sev = 'Warning'; Text = 'VM(s) run a guest OS past vendor end of support (heuristic match on the OS name).'; Rec = 'Check the VMware Guest OS Compatibility Guide for ESX 9.x and plan the OS upgrade or an isolation plan.' },
        @{ Key = 'consolidation'; Sev = 'Warning'; Text = 'VM(s) need snapshot consolidation.'; Rec = 'Consolidate disks now. This is a health problem regardless of the upgrade.' },
        @{ Key = 'Tools not installed'; Sev = 'Warning'; Text = 'powered-on VM(s) have no VMware Tools.'; Rec = 'Install Tools. Without them, the guest cannot be quiesced and some operations fail.' },
        @{ Key = 'Tools outdated'; Sev = 'Info'; Text = 'powered-on VM(s) have outdated VMware Tools.'; Rec = 'Update Tools ahead of the hardware version change.' }
    )
    foreach ($r in $vmRules) {
        $hit = @($vmList | Where-Object { $_.Flags -match [regex]::Escape($r.Key) })
        if ($hit.Count -gt 0) { Add-Finding $F $envName $r.Sev 'VM' (Sample $hit.VM) "$($hit.Count) $($r.Text)" $r.Rec }
    }
    $off = @($vmList | Where-Object { $_.Power -match 'Off' })
    if ($off.Count -gt 0) { Add-Finding $F $envName 'Info' 'VM' "$($off.Count) powered-off VMs" "$($off.Count) of $($vmList.Count) VMs are powered off." 'Ask the owner whether they are still needed. Fewer VMs means less to test and migrate.' }

    # ---- licences
    $lics = New-Object System.Collections.Generic.List[object]
    foreach ($l in $d.vLicense) {
        $key = "$(V $l @('key'))"; if ($key.Length -gt 5) { $key = $key.Substring(0, 5) + '-*****' }
        $exp = "$(V $l @('expirationdate'))"; $expDt = ToDate $exp
        $sev = 'OK'
        if ($expDt -and $expDt -lt $now.AddDays(90) -and $expDt.Year -lt 2100) { $sev = if ($expDt -lt $now) { 'Blocker' } else { 'Warning' } }
        $lics.Add([pscustomobject]@{ Environment = $envName; Name = "$(V $l @('name'))"; Key = $key; Total = "$(V $l @('total'))"; Used = "$(V $l @('used'))"; Expires = if ($expDt) { $expDt.ToString('yyyy-MM-dd') } else { $exp }; Severity = $sev })
        if ($sev -ne 'OK') { Add-Finding $F $envName $sev 'Licence' "$(V $l @('name'))" "Licence expires $($expDt.ToString('yyyy-MM-dd'))." 'Sort out the entitlement as part of the move to subscription licences.' }
    }
    Add-Finding $F $envName 'Info' 'Licence' $envName 'VCF and VVF 9 use subscription licence files, not 25-character keys.' 'Get the licence file from the Broadcom Support Portal before the upgrade window.'
    if ($t.perCoreLicensing) {
        $small = @($hosts | Where-Object { $_.Sockets -gt 0 -and $_.CoresPerCpu -lt 16 })
        if ($small.Count -gt 0) { Add-Finding $F $envName 'Info' 'Licence' (Sample $small.Host) "$($small.Count) host(s) have fewer than 16 cores per CPU." 'Broadcom licenses a minimum of 16 cores per CPU, so these hosts cost more per core than they show. Include them in consolidation planning.' }
    }

    # ---- sizing
    $pOn = @($vmList | Where-Object { $_.Power -match 'On' })
    $coresTot = ($hosts | Measure-Object Cores -Sum).Sum; $vcpuOn = ($pOn | Measure-Object vCPU -Sum).Sum
    $sizing = [ordered]@{
        'Hosts' = $hosts.Count; 'Clusters' = $clusters.Count; 'CPU sockets' = ($hosts | Measure-Object Sockets -Sum).Sum
        'Physical cores' = $coresTot; 'Licensed cores (min 16 per CPU)' = ($hosts | Measure-Object LicensedCores -Sum).Sum
        'Host memory (GiB)' = [math]::Round(($hosts | Measure-Object MemGiB -Sum).Sum, 0)
        'VMs (excluding templates)' = $vmList.Count; 'Powered-on VMs' = $pOn.Count; 'Powered-on vCPUs' = $vcpuOn
        'vCPU to core ratio' = if ($coresTot -gt 0) { [math]::Round($vcpuOn / $coresTot, 2) } else { 0 }
        'Powered-on VM memory (GiB)' = [math]::Round(($pOn | Measure-Object MemGiB -Sum).Sum, 0)
        'Datastore capacity (GiB)' = [math]::Round(($stores | Measure-Object CapacityGiB -Sum).Sum, 0)
        'Datastore free (GiB)' = [math]::Round(($stores | Measure-Object FreeGiB -Sum).Sum, 0)
    }
    $hwDist = @($vmList | Group-Object HwVersion | Sort-Object { [int]$_.Name } | ForEach-Object { [pscustomobject]@{ Environment = $envName; Item = "Hardware version $($_.Name)"; Count = $_.Count } })
    $esxDist = @($hosts | Group-Object EsxVersion | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Environment = $envName; Item = "ESXi $($_.Name)"; Count = $_.Count } })
    $osDist = @($vmList | Where-Object { $_.GuestOS } | Group-Object GuestOS | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { [pscustomobject]@{ Environment = $envName; Item = $_.Name; Count = $_.Count } })

    # ---- features
    $featRows = New-Object System.Collections.Generic.List[object]
    $vsanDs = @($stores | Where-Object { $_.Type -match 'vsan' }).Count
    $vsanCl = @($clusters | Where-Object { $_.vSAN -eq 'On' }).Count
    $nsxSeen = $false
    foreach ($n in $d.vNetwork) { if ("$(V $n @('network','portgroup','switch'))" -match 'nsx|segment|overlay') { $nsxSeen = $true; break } }
    $detect = @{
        'vSAN' = if ($vsanDs -gt 0 -or $vsanCl -gt 0) { "Detected: $vsanDs vSAN datastore(s), $vsanCl cluster(s) with vSAN on." } else { 'Not detected in the export.' }
        'NSX'  = if ($nsxSeen) { 'Possible NSX segments seen in network names (heuristic).' } else { 'Not detected in the export. RVTools cannot see NSX directly.' }
    }
    foreach ($fn in $features) {
        $items = $script:Cfg.features.$fn
        $det = if ($detect.ContainsKey($fn)) { $detect[$fn] } else { 'RVTools cannot see this. Confirm with the customer.' }
        foreach ($i in @($items)) { $featRows.Add([pscustomobject]@{ Environment = $envName; Feature = $fn; Detected = $det; Check = $i }) }
        if ($fn -eq 'NSX' -and $targetName -like 'VVF*') { Add-Finding $F $envName 'Warning' 'Feature' 'NSX' 'NSX was selected with a VVF target. VVF has a smaller entitlement than VCF.' 'Confirm NSX is covered by the VVF subscription, or target VCF instead.' }
        if ($fn -eq 'VCF Automation' -and $targetName -like 'VVF*') { Add-Finding $F $envName 'Warning' 'Feature' 'VCF Automation' 'VCF Automation was selected with a VVF target. VVF has a smaller entitlement than VCF.' 'Confirm it is covered by the VVF subscription, or target VCF instead.' }
    }
    if ($vsanDs -gt 0 -and $features -notcontains 'vSAN') { Add-Finding $F $envName 'Info' 'Feature' 'vSAN' "vSAN found ($vsanDs datastore(s)) but not in the selected features." 'Add vSAN to the feature list so the HCL and health checks are included.' }
    if ($nsxSeen -and $features -notcontains 'NSX') { Add-Finding $F $envName 'Info' 'Feature' 'NSX' 'Network names suggest NSX is in use but it is not in the selected features.' 'Confirm with the customer and add NSX to the feature list.' }

    # ---- HCL table (unique models)
    $hclRows = @($hosts | Group-Object Vendor, Model, CpuModel | ForEach-Object {
            $g = $_.Group[0]
            [pscustomobject]@{ Environment = $envName; Vendor = $g.Vendor; Model = $g.Model; CpuModel = $g.CpuModel; Hosts = $_.Count; CpuFamily = $g.CpuFamily; CpuStatus = $g.CpuStatus; HclStatus = $g.Hcl }
        })

    $blockers = @($F | Where-Object { $_.Severity -eq 'Blocker' }).Count
    $warns = @($F | Where-Object { $_.Severity -eq 'Warning' }).Count
    $readiness = if ($blockers -gt 0) { 'Blocked' } elseif ($warns -gt 0) { 'Ready with actions' } else { 'Ready' }
    $sorted = @($F | Sort-Object @{Expression = { switch ($_.Severity) { 'Blocker' { 0 } 'Warning' { 1 } default { 2 } } } }, Area, Object)

    [pscustomobject]@{
        Name = $envName; VCenters = $vc; Hosts = $hosts.ToArray(); Clusters = $clusters.ToArray(); Datastores = $stores.ToArray(); VMs = $vms.ToArray(); Licences = $lics.ToArray()
        Findings = $sorted; Sizing = $sizing; HwDist = $hwDist; EsxDist = $esxDist; OsDist = $osDist; FeatureRows = $featRows.ToArray(); Hcl = $hclRows
        Blockers = $blockers; Warnings = $warns; Infos = @($F | Where-Object { $_.Severity -eq 'Info' }).Count; Readiness = $readiness
        VCenterText = (($vc | ForEach-Object { $_.Version.Label }) -join ', '); EsxText = (($esxDist | ForEach-Object { $_.Item -replace '^ESXi ', '' }) -join ', ')
    }
}
#endregion

#region ---------- Excel report ----------
function To-Rows($objs, [string[]]$props) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($o in $objs) { $out.Add(@($props | ForEach-Object { $o.$_ })) }
    , $out
}

function Build-Workbook($envs, [string]$targetName, [string[]]$features, [string]$customer, [string]$file) {
    $used = New-Object System.Collections.Generic.List[string]
    $sheets = New-Object System.Collections.Generic.List[object]
    $name = { param($n) $u = SafeSheetName $n $used; $used.Add($u.ToLower()); $u }
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $ovName = & $name 'Overview'
    $envSheetNames = @{}
    foreach ($e in $envs) { $envSheetNames[$e.Name] = & $name ("Env - " + $e.Name) }

    # Overview
    $pre = @(
        @{ Text = "RVTools upgrade readiness$(if ($customer) { ' - ' + $customer })"; Style = 8 },
        @{ Text = "Target: $targetName    Features: $(if ($features) { $features -join ', ' } else { 'none selected' })    Generated: $stamp" },
        @{ Text = "Click an environment name to jump to its sheet. Blocker = will stop the upgrade. Warning = needs action or checking. Info = for awareness." },
        @{ Text = '' })
    $hdr = 'Environment', 'vCenter', 'ESXi versions', 'Hosts', 'VMs', 'Cores', 'Licensed cores', 'Blockers', 'Warnings', 'Readiness'
    $rows = New-Object System.Collections.Generic.List[object]; $links = @{}
    $i = 0
    foreach ($e in $envs) {
        $rows.Add(@($e.Name, $e.VCenterText, $e.EsxText, $e.Sizing['Hosts'], $e.Sizing['VMs (excluding templates)'], $e.Sizing['Physical cores'], $e.Sizing['Licensed cores (min 16 per CPU)'], $e.Blockers, $e.Warnings, $e.Readiness))
        $links["A$($pre.Count + 2 + $i)"] = $envSheetNames[$e.Name]; $i++
    }
    $sheets.Add(@{ Name = $ovName; PreRows = $pre; HeaderRow = $pre.Count + 1; Headers = $hdr; Rows = $rows; Widths = @(34, 16, 28, 8, 8, 8, 14, 10, 10, 20); SevCols = @(9); LinkCols = @(0); LinkTargets = $links; Freeze = $false; NoFilter = $true })

    # Per-environment sheets
    foreach ($e in $envs) {
        $pre = New-Object System.Collections.Generic.List[object]
        $pre.Add(@{ Text = $e.Name; Style = 8 })
        $pre.Add(@{ Text = "Readiness for ${targetName}: $($e.Readiness)   (Blockers $($e.Blockers), Warnings $($e.Warnings), Info $($e.Infos))" })
        $pre.Add(@{ Text = "vCenter: $($e.VCenterText)    ESXi: $($e.EsxText)" })
        foreach ($k in $e.Sizing.Keys) { $pre.Add(@{ Text = "$k`: $($e.Sizing[$k])" }) }
        $pre.Add(@{ Text = '' })
        $sheets.Add(@{ Name = $envSheetNames[$e.Name]; PreRows = $pre; HeaderRow = $pre.Count + 1; Freeze = $false
                Headers = @('Severity', 'Area', 'Object', 'Finding', 'Recommendation'); Widths = @(11, 12, 32, 70, 70); SevCols = @(0); WrapCols = @(2, 3, 4)
                Rows = (To-Rows $e.Findings @('Severity', 'Area', 'Object', 'Finding', 'Recommendation')) })
    }

    $all = { param($prop) @($envs | ForEach-Object { $_.$prop }) }
    $sheets.Add(@{ Name = (& $name 'Findings'); Headers = @('Environment', 'Severity', 'Area', 'Object', 'Finding', 'Recommendation'); Widths = @(26, 11, 12, 32, 70, 70); SevCols = @(1); WrapCols = @(3, 4, 5)
            Rows = (To-Rows (& $all 'Findings') @('Environment', 'Severity', 'Area', 'Object', 'Finding', 'Recommendation')) })
    $hp = 'Environment', 'Host', 'Cluster', 'Vendor', 'Model', 'CpuModel', 'Sockets', 'CoresPerCpu', 'Cores', 'LicensedCores', 'MemGiB', 'VMs', 'EsxVersion', 'CpuFamily', 'CpuStatus', 'Hcl', 'Severity', 'Notes'
    $sheets.Add(@{ Name = (& $name 'Hosts'); Headers = $hp; Widths = @(26, 30, 22, 14, 26, 44, 8, 8, 8, 10, 9, 6, 12, 30, 12, 12, 10, 60); SevCols = @(16); Rows = (To-Rows (& $all 'Hosts') $hp) })
    $cp = 'Environment', 'Cluster', 'Hosts', 'VMs', 'Cores', 'MemGiB', 'HA', 'DRS', 'vSAN', 'EsxVersions', 'BlockedHosts', 'Severity'
    $sheets.Add(@{ Name = (& $name 'Clusters'); Headers = $cp; Widths = @(26, 28, 7, 7, 8, 9, 6, 6, 6, 24, 12, 10); SevCols = @(11); Rows = (To-Rows (& $all 'Clusters') $cp) })
    $vp = 'Environment', 'VM', 'Power', 'Cluster', 'Host', 'vCPU', 'MemGiB', 'HwVersion', 'Tools', 'GuestOS', 'Snapshots', 'Severity', 'Flags'
    $vmRows = @($envs | ForEach-Object { $_.VMs | Where-Object { -not $_.Template } })
    $sheets.Add(@{ Name = (& $name 'VMs'); Headers = $vp; Widths = @(26, 34, 11, 22, 30, 6, 8, 9, 20, 36, 9, 10, 60); SevCols = @(11); Rows = (To-Rows $vmRows $vp) })
    $dp = 'Environment', 'Datastore', 'Type', 'CapacityGiB', 'FreeGiB', 'UsedPct', 'VmfsVersion', 'Severity'
    $sheets.Add(@{ Name = (& $name 'Datastores'); Headers = $dp; Widths = @(26, 34, 10, 12, 10, 9, 12, 10); SevCols = @(7); Rows = (To-Rows (& $all 'Datastores') $dp) })
    $lp = 'Environment', 'Name', 'Key', 'Total', 'Used', 'Expires', 'Severity'
    $sheets.Add(@{ Name = (& $name 'Licences'); Headers = $lp; Widths = @(26, 40, 14, 10, 10, 14, 10); SevCols = @(6); Rows = (To-Rows (& $all 'Licences') $lp) })
    $hc = 'Environment', 'Vendor', 'Model', 'CpuModel', 'Hosts', 'CpuFamily', 'CpuStatus', 'HclStatus'
    $sheets.Add(@{ Name = (& $name 'HCL'); Headers = $hc; Widths = @(26, 14, 30, 44, 7, 32, 12, 14); SevCols = @(6); Rows = (To-Rows (& $all 'Hcl') $hc)
            PreRows = @(@{ Text = 'Search each Vendor + Model on https://compatibilityguide.broadcom.com/ for the target release. Import a Broadcom export with -ImportHcl to fill HclStatus.' }); HeaderRow = 2; Freeze = $false })
    $fp = 'Environment', 'Feature', 'Detected', 'Check'
    $sheets.Add(@{ Name = (& $name 'Features'); Headers = $fp; Widths = @(26, 22, 50, 90); WrapCols = @(2, 3); Rows = (To-Rows (& $all 'FeatureRows') $fp) })
    $dist = New-Object System.Collections.Generic.List[object]
    foreach ($e in $envs) { foreach ($x in @($e.EsxDist) + @($e.HwDist) + @($e.OsDist)) { $dist.Add(@($x.Environment, $x.Item, $x.Count)) } }
    $sheets.Add(@{ Name = (& $name 'Distributions'); Headers = @('Environment', 'Item', 'Count'); Widths = @(26, 60, 8); Rows = $dist })
    $tn = $script:Cfg.targets.$targetName
    $notes = New-Object System.Collections.Generic.List[object]
    foreach ($n in $tn.notes) { $notes.Add(@('Target', $n)) }
    foreach ($n in $script:Cfg.manualChecks) { $notes.Add(@('Manual check', $n)) }
    foreach ($n in $script:Cfg.sources) { $notes.Add(@('Source', "$($n.title) - $($n.url)")) }
    $notes.Add(@('Rules', "Bundled rules last checked $($script:Cfg.rulesChecked). Broadcom changes these. Re-check before customer use."))
    $sheets.Add(@{ Name = (& $name 'Notes and sources'); Headers = @('Type', 'Note'); Widths = @(16, 130); WrapCols = @(1); Rows = $notes })
    Write-Xlsx $file $sheets
}
#endregion

#region ---------- HTML report ----------
function Html-Table($rows, [string[]]$cols, [string[]]$heads, [string[]]$sevCols) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tw"><table class="dt"><thead><tr>')
    for ($i = 0; $i -lt $cols.Count; $i++) { [void]$sb.Append("<th>$(Enc $heads[$i])</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $rows) {
        $sevVal = $r.Severity
        [void]$sb.Append("<tr data-sev=`"$(Enc $sevVal)`">")
        foreach ($c in $cols) {
            $v = "$($r.$c)"
            if ($sevCols -contains $c) { [void]$sb.Append("<td><span class=`"pill $(($v -replace '\s', '').ToLower())`">$(Enc $v)</span></td>") }
            else { [void]$sb.Append("<td>$(Enc $v)</td>") }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    $sb.ToString()
}

function Build-Html($envs, [string]$targetName, [string[]]$features, [string]$customer, [string]$file) {
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $sb = New-Object System.Text.StringBuilder
    $title = "Upgrade readiness for $targetName$(if ($customer) { ' - ' + $customer })"
    [void]$sb.Append(@"
<!doctype html><html lang="en-GB"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>$(Enc $title)</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c2430;--mute:#5b6675;--line:#dfe3e8;--acc:#1f3864;--red:#c0392b;--redbg:#fde3e0;--amb:#9a6700;--ambbg:#fff1c9;--grn:#1e7e34;--grnbg:#dcf3e2;--blu:#1f5a99;--blubg:#e1edf9}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--ink);display:flex;min-height:100vh}
nav{width:250px;background:var(--acc);color:#fff;padding:16px 0;position:sticky;top:0;height:100vh;overflow:auto;flex:none}
nav h1{font-size:15px;margin:0 16px 4px}nav small{display:block;margin:0 16px 14px;opacity:.75}
nav a{display:block;color:#dbe4f5;text-decoration:none;padding:8px 16px;border-left:3px solid transparent;cursor:pointer}
nav a:hover{background:rgba(255,255,255,.08)}nav a.on{background:rgba(255,255,255,.14);border-left-color:#fff;color:#fff}
nav .grp{margin:14px 16px 4px;font-size:11px;letter-spacing:.08em;text-transform:uppercase;opacity:.6}
nav a .dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:8px}
.dot.blocked{background:#ff8b7f}.dot.readywithactions{background:#ffd166}.dot.ready{background:#7ddc95}
main{flex:1;padding:24px 28px;min-width:0}section{display:none}section.on{display:block}
h2{margin:0 0 4px;font-size:22px}h3{margin:26px 0 8px;font-size:16px}.sub{color:var(--mute);margin-bottom:16px}
.kpis{display:flex;flex-wrap:wrap;gap:10px;margin:14px 0}.kpi{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 14px;min-width:120px}
.kpi b{display:block;font-size:22px}.kpi span{color:var(--mute);font-size:12px}
.pill{display:inline-block;padding:1px 9px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap}
.pill.blocker,.pill.blocked{background:var(--redbg);color:var(--red)}.pill.warning,.pill.readywithactions{background:var(--ambbg);color:var(--amb)}
.pill.ok,.pill.ready,.pill.listed{background:var(--grnbg);color:var(--grn)}.pill.info,.pill.verify,.pill.notchecked{background:var(--blubg);color:var(--blu)}.pill.notlisted{background:var(--ambbg);color:var(--amb)}
.tw{overflow:auto;background:var(--card);border:1px solid var(--line);border-radius:8px;max-height:560px}
table.dt{border-collapse:collapse;width:100%}table.dt th{position:sticky;top:0;background:#eef1f5;text-align:left;padding:7px 10px;font-size:12px;border-bottom:1px solid var(--line);white-space:nowrap}
table.dt td{padding:6px 10px;border-bottom:1px solid #eef0f3;vertical-align:top}table.dt tr:hover td{background:#fafbfc}
.tools{display:flex;gap:8px;margin:0 0 8px;flex-wrap:wrap;align-items:center}
.tools button{border:1px solid var(--line);background:#fff;padding:4px 12px;border-radius:14px;cursor:pointer;font-size:12px}.tools button.on{background:var(--acc);color:#fff;border-color:var(--acc)}
.tools input{border:1px solid var(--line);border-radius:6px;padding:5px 9px;min-width:220px}
.banner{padding:12px 16px;border-radius:8px;font-weight:600;margin:10px 0}.banner.blocked{background:var(--redbg);color:var(--red)}.banner.readywithactions{background:var(--ambbg);color:var(--amb)}.banner.ready{background:var(--grnbg);color:var(--grn)}
ul.n{margin:6px 0 0 18px;padding:0}ul.n li{margin:4px 0}a.ext{color:var(--blu)}
@media(max-width:800px){body{display:block}nav{width:auto;height:auto;position:static}}
@media print{nav{display:none}section{display:block!important;page-break-after:always}.tw{max-height:none}}
</style></head><body>
"@)
    # nav
    [void]$sb.Append("<nav><h1>$(Enc $targetName)</h1><small>$(if ($customer) { (Enc $customer) + '<br>' })$stamp</small><a class=`"on`" data-t=`"overview`">Overview</a><div class=`"grp`">Environments</div>")
    for ($i = 0; $i -lt $envs.Count; $i++) {
        [void]$sb.Append("<a data-t=`"env$i`"><span class=`"dot $(($envs[$i].Readiness -replace '\s','').ToLower())`"></span>$(Enc $envs[$i].Name)</a>")
    }
    [void]$sb.Append('<div class="grp">Reference</div><a data-t="notes">Notes and sources</a></nav><main>')

    # overview
    [void]$sb.Append("<section id=`"overview`" class=`"on`"><h2>$(Enc $title)</h2><div class=`"sub`">Generated $stamp from RVTools exports. Features: $(if ($features) { Enc ($features -join ', ') } else { 'none selected' }).</div>")
    $tb = ($envs | Measure-Object Blockers -Sum).Sum; $tw = ($envs | Measure-Object Warnings -Sum).Sum
    [void]$sb.Append("<div class=`"kpis`"><div class=`"kpi`"><b>$($envs.Count)</b><span>environments</span></div><div class=`"kpi`"><b>$(($envs | ForEach-Object { $_.Sizing['Hosts'] } | Measure-Object -Sum).Sum)</b><span>hosts</span></div><div class=`"kpi`"><b>$(($envs | ForEach-Object { $_.Sizing['VMs (excluding templates)'] } | Measure-Object -Sum).Sum)</b><span>VMs</span></div><div class=`"kpi`"><b>$(($envs | ForEach-Object { $_.Sizing['Licensed cores (min 16 per CPU)'] } | Measure-Object -Sum).Sum)</b><span>licensed cores</span></div><div class=`"kpi`"><b style=`"color:var(--red)`">$tb</b><span>blockers</span></div><div class=`"kpi`"><b style=`"color:var(--amb)`">$tw</b><span>warnings</span></div></div>")
    [void]$sb.Append('<div class="tw"><table class="dt"><thead><tr><th>Environment</th><th>vCenter</th><th>ESXi versions</th><th>Hosts</th><th>VMs</th><th>Cores</th><th>Licensed cores</th><th>Blockers</th><th>Warnings</th><th>Readiness</th></tr></thead><tbody>')
    for ($i = 0; $i -lt $envs.Count; $i++) {
        $e = $envs[$i]
        [void]$sb.Append("<tr><td><a class=`"ext`" href=`"#env$i`" data-go=`"env$i`">$(Enc $e.Name)</a></td><td>$(Enc $e.VCenterText)</td><td>$(Enc $e.EsxText)</td><td>$($e.Sizing['Hosts'])</td><td>$($e.Sizing['VMs (excluding templates)'])</td><td>$($e.Sizing['Physical cores'])</td><td>$($e.Sizing['Licensed cores (min 16 per CPU)'])</td><td>$($e.Blockers)</td><td>$($e.Warnings)</td><td><span class=`"pill $(($e.Readiness -replace '\s','').ToLower())`">$(Enc $e.Readiness)</span></td></tr>")
    }
    [void]$sb.Append('</tbody></table></div>')
    [void]$sb.Append('<h3>About this target</h3><ul class="n">')
    foreach ($n in $script:Cfg.targets.$targetName.notes) { [void]$sb.Append("<li>$(Enc $n)</li>") }
    [void]$sb.Append('</ul></section>')

    # env sections
    for ($i = 0; $i -lt $envs.Count; $i++) {
        $e = $envs[$i]; $rc = ($e.Readiness -replace '\s', '').ToLower()
        [void]$sb.Append("<section id=`"env$i`"><h2>$(Enc $e.Name)</h2><div class=`"sub`">vCenter $(Enc $e.VCenterText). ESXi $(Enc $e.EsxText).</div>")
        [void]$sb.Append("<div class=`"banner $rc`">$(Enc $e.Readiness) for $(Enc $targetName): $($e.Blockers) blocker(s), $($e.Warnings) warning(s), $($e.Infos) info.</div>")
        [void]$sb.Append('<div class="kpis">')
        foreach ($k in @('Hosts', 'Physical cores', 'Licensed cores (min 16 per CPU)', 'VMs (excluding templates)', 'Powered-on VMs', 'vCPU to core ratio', 'Host memory (GiB)', 'Datastore capacity (GiB)')) { [void]$sb.Append("<div class=`"kpi`"><b>$($e.Sizing[$k])</b><span>$(Enc $k)</span></div>") }
        [void]$sb.Append('</div>')
        [void]$sb.Append('<h3>Findings</h3><div class="tools" data-tbl="f' + $i + '"><button class="on" data-f="">All</button><button data-f="Blocker">Blockers</button><button data-f="Warning">Warnings</button><button data-f="Info">Info</button><input placeholder="Filter text" type="search"></div><div id="f' + $i + '">')
        [void]$sb.Append((Html-Table $e.Findings @('Severity', 'Area', 'Object', 'Finding', 'Recommendation') @('Severity', 'Area', 'Object', 'Finding', 'Recommendation') @('Severity')))
        [void]$sb.Append('</div>')
        [void]$sb.Append('<h3>Hosts</h3><div class="tools" data-tbl="h' + $i + '"><input placeholder="Filter hosts" type="search"></div><div id="h' + $i + '">')
        [void]$sb.Append((Html-Table $e.Hosts @('Host', 'Cluster', 'Model', 'CpuModel', 'Cores', 'MemGiB', 'EsxVersion', 'CpuFamily', 'CpuStatus', 'Hcl', 'Severity', 'Notes') @('Host', 'Cluster', 'Model', 'CPU', 'Cores', 'Mem GiB', 'ESXi', 'CPU family', 'CPU status', 'HCL', 'Severity', 'Notes') @('Severity', 'CpuStatus', 'Hcl')))
        [void]$sb.Append('</div><h3>Clusters</h3>')
        [void]$sb.Append((Html-Table $e.Clusters @('Cluster', 'Hosts', 'VMs', 'Cores', 'MemGiB', 'HA', 'DRS', 'vSAN', 'EsxVersions', 'Severity') @('Cluster', 'Hosts', 'VMs', 'Cores', 'Mem GiB', 'HA', 'DRS', 'vSAN', 'ESXi', 'Severity') @('Severity')))
        if ($e.FeatureRows.Count -gt 0) {
            [void]$sb.Append('<h3>Feature checks</h3>')
            [void]$sb.Append((Html-Table $e.FeatureRows @('Feature', 'Detected', 'Check') @('Feature', 'Detected in RVTools', 'Check') @()))
        }
        $flag = @($e.VMs | Where-Object { -not $_.Template -and $_.Severity -ne 'OK' } | Sort-Object @{Expression = { if ($_.Severity -eq 'Warning') { 0 } else { 1 } } }, VM)
        [void]$sb.Append("<h3>VMs needing attention ($($flag.Count))</h3><div class=`"tools`" data-tbl=`"v$i`"><input placeholder=`"Filter VMs`" type=`"search`"></div><div id=`"v$i`">")
        [void]$sb.Append((Html-Table ($flag | Select-Object -First 500) @('VM', 'Power', 'Cluster', 'HwVersion', 'Tools', 'GuestOS', 'Snapshots', 'Severity', 'Flags') @('VM', 'Power', 'Cluster', 'HW', 'Tools', 'Guest OS', 'Snaps', 'Severity', 'Flags') @('Severity')))
        if ($flag.Count -gt 500) { [void]$sb.Append('<div class="sub">Showing the first 500. The Excel workbook has all VMs.</div>') }
        [void]$sb.Append('</div><h3>Datastores</h3>')
        [void]$sb.Append((Html-Table $e.Datastores @('Datastore', 'Type', 'CapacityGiB', 'FreeGiB', 'UsedPct', 'VmfsVersion', 'Severity') @('Datastore', 'Type', 'Capacity GiB', 'Free GiB', 'Used %', 'VMFS', 'Severity') @('Severity')))
        [void]$sb.Append('<h3>Hardware compatibility</h3><div class="sub">Search each vendor and model on the <a class="ext" target="_blank" rel="noopener" href="https://compatibilityguide.broadcom.com/">Broadcom Compatibility Guide</a> for the target release. HCL status is filled only if HCL data has been imported.</div>')
        [void]$sb.Append((Html-Table $e.Hcl @('Vendor', 'Model', 'CpuModel', 'Hosts', 'CpuFamily', 'CpuStatus', 'HclStatus') @('Vendor', 'Model', 'CPU', 'Hosts', 'CPU family', 'CPU status', 'HCL') @('CpuStatus', 'HclStatus')))
        if ($e.Licences.Count -gt 0) { [void]$sb.Append('<h3>Licences</h3>'); [void]$sb.Append((Html-Table $e.Licences @('Name', 'Key', 'Total', 'Used', 'Expires', 'Severity') @('Name', 'Key', 'Total', 'Used', 'Expires', 'Severity') @('Severity'))) }
        [void]$sb.Append('</section>')
    }
    # notes
    [void]$sb.Append("<section id=`"notes`"><h2>Notes and sources</h2><div class=`"sub`">Bundled rules last checked $(Enc $script:Cfg.rulesChecked). Broadcom changes these, so re-check before customer use.</div><h3>Manual checks RVTools cannot answer</h3><ul class=`"n`">")
    foreach ($n in $script:Cfg.manualChecks) { [void]$sb.Append("<li>$(Enc $n)</li>") }
    [void]$sb.Append('</ul><h3>Sources</h3><ul class="n">')
    foreach ($s in $script:Cfg.sources) { [void]$sb.Append("<li><a class=`"ext`" target=`"_blank`" rel=`"noopener`" href=`"$(Enc $s.url)`">$(Enc $s.title)</a></li>") }
    [void]$sb.Append('</ul></section></main>')
    [void]$sb.Append(@'
<script>
(function(){
var links=document.querySelectorAll('nav a[data-t]');
function show(id){document.querySelectorAll('section').forEach(function(s){s.classList.toggle('on',s.id===id)});links.forEach(function(a){a.classList.toggle('on',a.dataset.t===id)});window.scrollTo(0,0);try{history.replaceState(null,'','#'+id)}catch(e){}}
links.forEach(function(a){a.addEventListener('click',function(){show(a.dataset.t)})});
document.querySelectorAll('[data-go]').forEach(function(a){a.addEventListener('click',function(ev){ev.preventDefault();show(a.dataset.go)})});
document.querySelectorAll('.tools').forEach(function(t){
 var box=document.getElementById(t.dataset.tbl);if(!box)return;var sev='',txt='';
 function apply(){box.querySelectorAll('tbody tr').forEach(function(r){var ok=(!sev||r.dataset.sev===sev)&&(!txt||r.textContent.toLowerCase().indexOf(txt)>-1);r.style.display=ok?'':'none'})}
 t.querySelectorAll('button').forEach(function(b){b.addEventListener('click',function(){t.querySelectorAll('button').forEach(function(x){x.classList.remove('on')});b.classList.add('on');sev=b.dataset.f;apply()})});
 var i=t.querySelector('input');if(i)i.addEventListener('input',function(){txt=i.value.toLowerCase();apply()});
});
var h=location.hash.replace('#','');if(h&&document.getElementById(h))show(h);
})();
</script></body></html>
'@)
    [System.IO.File]::WriteAllText($file, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
}
#endregion

#region ---------- main ----------
function Select-One([string]$prompt, [string[]]$options, [string]$default) {
    Write-Host ''; Write-Host $prompt
    for ($i = 0; $i -lt $options.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i + 1), $options[$i]) }
    $a = Read-Host "Choose 1-$($options.Count) [default: $default]"
    if (-not $a) { return $default }
    $n = 0; if ([int]::TryParse($a, [ref]$n) -and $n -ge 1 -and $n -le $options.Count) { return $options[$n - 1] }
    return $default
}
function Select-Many([string]$prompt, [string[]]$options) {
    Write-Host ''; Write-Host $prompt
    for ($i = 0; $i -lt $options.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i + 1), $options[$i]) }
    $a = Read-Host 'Numbers separated by commas, or Enter for none'
    if (-not $a) { return @() }
    $out = @(); foreach ($x in ($a -split '[,\s]+')) { $n = 0; if ([int]::TryParse($x, [ref]$n) -and $n -ge 1 -and $n -le $options.Count) { $out += $options[$n - 1] } }
    $out
}

Load-Config
if ($ImportHcl) { Import-HclExport $ImportHcl $HclTarget; if (-not $Path) { return } }

if (-not $Path) {
    if ($Interactive) { $p = Read-Host 'Path to RVTools .xlsx file(s) or a folder'; $Path = @($p.Trim('"')) }
    else { throw 'Give -Path with one or more RVTools files or folders (or use -Interactive).' }
}
if ($Interactive) {
    if (-not $Target) { $Target = Select-One 'Upgrade target:' @($script:Cfg.targets.PSObject.Properties.Name) 'VCF 9.1' }
    if (-not $Features) { $Features = Select-Many 'Features to include:' @($script:Cfg.features.PSObject.Properties.Name) }
    if (-not $Customer) { $Customer = Read-Host 'Customer name (optional)' }
}
if (-not $Target) { $Target = 'VCF 9.1' }
if (-not $Features) { $Features = @() }
foreach ($f in $Features) { if (-not $script:Cfg.features.PSObject.Properties[$f]) { throw "Unknown feature '$f'. Valid: $($script:Cfg.features.PSObject.Properties.Name -join ', ')" } }
$tcfg = $script:Cfg.targets.$Target

# resolve input files
$xlsx = New-Object System.Collections.Generic.List[string]; $csvGroups = @{}
foreach ($p in $Path) {
    foreach ($item in @(Get-Item -Path $p -ErrorAction SilentlyContinue)) {
        if ($item.PSIsContainer) {
            Get-ChildItem $item.FullName -Filter *.xlsx -File | Where-Object { $_.Name -notlike '~$*' } | ForEach-Object { $xlsx.Add($_.FullName) }
            $csv = @(Get-ChildItem $item.FullName -Filter '*tab*.csv' -File)
            if ($csv.Count -gt 0) { $csvGroups[$item.Name] = $csv.FullName }
        }
        elseif ($item.Extension -eq '.xlsx') { $xlsx.Add($item.FullName) }
        elseif ($item.Extension -eq '.csv') { $g = ($item.Name -replace '_?tab\w+\.csv$', ''); if (-not $csvGroups[$g]) { $csvGroups[$g] = @() }; $csvGroups[$g] += $item.FullName }
    }
}
if ($xlsx.Count -eq 0 -and $csvGroups.Count -eq 0) { throw 'No RVTools .xlsx or .csv files found in -Path.' }

$sheetNames = 'vInfo', 'vHost', 'vCluster', 'vDatastore', 'vSnapshot', 'vTools', 'vDisk', 'vLicense', 'vSource', 'vNetwork'
$envData = [ordered]@{}
$ingest = {
    param($wb, $fallback)
    foreach ($sn in $sheetNames) {
        foreach ($row in (Get-Sheet $wb $sn)) {
            $key = "$(V $row @('visdkserver'))"; if (-not $key) { $key = $fallback }
            if (-not $envData.Contains($key)) { $envData[$key] = @{}; foreach ($x in $sheetNames) { $envData[$key][$x] = New-Object System.Collections.Generic.List[hashtable] } }
            $envData[$key][$sn].Add($row)
        }
    }
}
foreach ($f in $xlsx) {
    Write-Host "Reading $([IO.Path]::GetFileName($f)) ..."
    $wb = Read-XlsxWorkbook $f
    if (-not ($wb.Keys | Where-Object { $_ -ieq 'vHost' -or $_ -ieq 'vInfo' })) { Write-Warning "$f has no vInfo or vHost tab. Skipped."; continue }
    & $ingest $wb ([IO.Path]::GetFileNameWithoutExtension($f))
}
foreach ($g in $csvGroups.Keys) { Write-Host "Reading CSV set $g ..."; & $ingest (Read-CsvSet $csvGroups[$g]) $g }
if ($envData.Count -eq 0) { throw 'Nothing readable was found in the input.' }

$hcl = @(Get-HclModels $Target)
if ($hcl.Count -eq 0) { Write-Host 'No HCL data loaded for this target. Hosts will be marked Not checked. See -ImportHcl.' }
$envs = @()
foreach ($k in $envData.Keys) {
    Write-Host "Analysing $k ($($envData[$k].vHost.Count) hosts, $($envData[$k].vInfo.Count) VMs) ..."
    $envs += Analyze-Environment $k $envData[$k] $tcfg $Target $Features $hcl
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$base = 'RVTools-Upgrade-Report-' + ($Target -replace '\s', '') + '-' + (Get-Date -Format 'yyyyMMdd')
$xl = Join-Path (Resolve-Path $OutputPath) "$base.xlsx"
$ht = Join-Path (Resolve-Path $OutputPath) "$base.html"
Build-Workbook $envs $Target $Features $Customer $xl
Build-Html $envs $Target $Features $Customer $ht
Write-Host ''
Write-Host "Done. $($envs.Count) environment(s)."
foreach ($e in $envs) { Write-Host ("  {0,-40} {1,-20} blockers {2}, warnings {3}" -f $e.Name, $e.Readiness, $e.Blockers, $e.Warnings) }
Write-Host "HTML : $ht"
Write-Host "Excel: $xl"
#endregion
