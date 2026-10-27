<#
.SYNOPSIS
    Parse a VCF SDDC Manager SoS --health-check text report into an Excel workbook.

.DESCRIPTION
    Assumes the report contains ASCII box-drawn tables, which is the standard
    SoS --health-check layout, e.g.:

    +-----+-------------------------------------+------------+-------+
    | SL# | Area                                | Title      | State |
    +-----+-------------------------------------+------------+-------+
    | 1   | ESXi : esxi-1.vrack.vsphere.local    | ESX Time   | GREEN |
    +-----+-------------------------------------+------------+-------+

    Every table in the file becomes its own worksheet, named after the line
    just above it (e.g. "NTP", "Password Check"). The State/Status column is
    colour-coded green/amber/red. A Summary sheet at the front totals the
    counts across the whole report.

    If your report doesn't use this box-table layout, send a snippet and the
    parser can be adjusted to match.

.PARAMETER InputFile
    Path to the SoS health-check text log.

.PARAMETER OutputFile
    Path to the .xlsx file to write. Defaults to the input file name with
    an .xlsx extension.

.EXAMPLE
    .\Sos-Report-To-Excel.ps1 -InputFile .\sos-health-check.log
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [string]$OutputFile
)

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "This script needs the ImportExcel module. Install it once with:"
    Write-Host "    Install-Module ImportExcel -Scope CurrentUser"
    exit 1
}
Import-Module ImportExcel

if (-not $OutputFile) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, "xlsx")
}
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}

$lines = Get-Content -Path $InputFile

$borderPattern = '^\+[-+]+\+\s*$'
$rowPattern    = '^\|.*\|\s*$'

function Split-Row {
    param([string]$Line)
    $trimmed = $Line.Trim().Trim('|')
    return @($trimmed -split '\|' | ForEach-Object { $_.Trim() })
}

function Find-Tables {
    param([string[]]$Lines)

    $tables = New-Object System.Collections.Generic.List[object]
    $n = $Lines.Count
    $i = 0

    while ($i -lt $n) {
        if ($Lines[$i] -match $borderPattern) {
            $start = $i
            $j = $i + 1
            if ($j -lt $n -and $Lines[$j] -match $rowPattern) {
                $header = Split-Row $Lines[$j]
                $j++
                if ($j -lt $n -and $Lines[$j] -match $borderPattern) {
                    $j++
                    $dataRows = New-Object System.Collections.Generic.List[object]
                    while ($j -lt $n -and $Lines[$j] -match $rowPattern) {
                        $dataRows.Add((Split-Row $Lines[$j]))
                        $j++
                    }
                    if ($j -lt $n -and $Lines[$j] -match $borderPattern) {
                        $j++
                    }

                    $title = $null
                    $lookBack = [Math]::Max($start - 6, 0)
                    for ($k = $start - 1; $k -ge $lookBack; $k--) {
                        $candidate = $Lines[$k].Trim()
                        if ($candidate -and ($candidate -notmatch $borderPattern) -and ($candidate -notmatch $rowPattern)) {
                            $title = $candidate
                            break
                        }
                    }

                    $tables.Add([PSCustomObject]@{
                        Title  = $title
                        Header = $header
                        Rows   = $dataRows
                    })
                    $i = $j
                    continue
                }
            }
        }
        $i++
    }
    return $tables
}

$tables = Find-Tables -Lines $lines

if ($tables.Count -eq 0) {
    Write-Host "No ASCII tables of the expected +---+ | ... | format were found."
    Write-Host "Send a snippet of the actual report and the parser can be adjusted."
    exit 1
}

$statusColours = [ordered]@{
    'GREEN'   = '#C6EFCE'; 'PASS'    = '#C6EFCE'; 'PASSED'  = '#C6EFCE'
    'YELLOW'  = '#FFEB9C'; 'WARN'    = '#FFEB9C'; 'WARNING' = '#FFEB9C'
    'RED'     = '#FFC7CE'; 'FAIL'    = '#FFC7CE'; 'FAILED'  = '#FFC7CE'
}

$summaryCounts  = @{}
$usedSheetNames = @{}

foreach ($table in $tables) {
    $baseName = if ($table.Title) { $table.Title } else { "Health Check" }
    $baseName = $baseName.Substring(0, [Math]::Min(31, $baseName.Length))
    $sheetName = $baseName
    $suffix = 2
    while ($usedSheetNames.ContainsKey($sheetName)) {
        $trimmedBase = $baseName.Substring(0, [Math]::Min(28, $baseName.Length))
        $sheetName = "$trimmedBase`_$suffix"
        $suffix++
    }
    $usedSheetNames[$sheetName] = $true

    $header = $table.Header
    $statusColIndex = -1
    for ($idx = 0; $idx -lt $header.Count; $idx++) {
        if ($header[$idx].Trim().ToLower() -in @('state', 'status')) {
            $statusColIndex = $idx
        }
    }

    $objects = foreach ($row in $table.Rows) {
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $header.Count; $c++) {
            $value = if ($c -lt $row.Count) { $row[$c] } else { '' }
            $obj[$header[$c]] = $value
            if ($c -eq $statusColIndex) {
                $key = $value.Trim().ToUpper()
                if ($summaryCounts.ContainsKey($key)) { $summaryCounts[$key]++ } else { $summaryCounts[$key] = 1 }
            }
        }
        [PSCustomObject]$obj
    }

    $conditionalRules = @()
    if ($statusColIndex -ge 0) {
        foreach ($statusValue in $statusColours.Keys) {
            $conditionalRules += New-ConditionalText -Text $statusValue -BackgroundColor $statusColours[$statusValue]
        }
    }

    $exportParams = @{
        Path          = $OutputFile
        WorksheetName = $sheetName
        AutoSize      = $true
        BoldTopRow    = $true
    }
    if ($conditionalRules.Count -gt 0) {
        $exportParams['ConditionalText'] = $conditionalRules
    }

    $objects | Export-Excel @exportParams
}

# Summary sheet
$summaryObjects = $summaryCounts.GetEnumerator() |
    Sort-Object -Property Value -Descending |
    ForEach-Object { [PSCustomObject]@{ Status = $_.Key; Count = $_.Value } }

$summaryConditional = $statusColours.Keys | ForEach-Object {
    New-ConditionalText -Text $_ -BackgroundColor $statusColours[$_]
}

$summaryObjects | Export-Excel -Path $OutputFile -WorksheetName "Summary" -AutoSize -BoldTopRow -ConditionalText $summaryConditional -MoveToStart

Write-Host "Wrote $($tables.Count) table(s) to $OutputFile"
