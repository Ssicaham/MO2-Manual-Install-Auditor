# =====================================================================
#  Find-ManualMods.ps1  (version 2.0)
# ---------------------------------------------------------------------
#  This script analyzes all mod folders in Mod Organizer 2 and
#  classifies them by the installationFile field in meta.ini,
#  and also collects additional data:
#    - presence of plugins (.esp / .esm / .esl) and their count
#    - presence of BSA / BA2 archives and their total size
#    - Nexus mod ID, version, Nexus page name (parsed from archive name)
#    - ready-to-click link to the Nexus page
#
#  Generates three reports:
#    1. ManualMods-Report.txt  - readable plain text report
#    2. ManualMods-Report.csv  - spreadsheet for Excel / filtering
#    3. ManualMods-Report.html - HTML with clickable links
#
#  How to run:
#    1. Open PowerShell in the script's folder.
#    2. If needed: Unblock-File .\Find-ManualMods.ps1
#    3. Run: .\Find-ManualMods.ps1
# =====================================================================


# ====================== EDIT THIS LINE ===============================
# Path to the "mods" folder of your MO2 instance.
# Examples:
#   "D:\MO2\Skyrim Special Edition\mods"
#   "D:\Skyrimmods\mods"
# =====================================================================
$modsPath = "D:\Skyrimmods\mods"


# ====================== GAME FOR NEXUS LINKS =========================
# Game slug used to build links of the form:
#   https://www.nexusmods.com/<gameSlug>/mods/<id>
# Allowed values: skyrimspecialedition, skyrim, fallout4,
# falloutnewvegas, oblivion, morrowind, fallout3, starfield, etc.
# =====================================================================
$nexusGameSlug = "skyrimspecialedition"


# ====================== ENABLE / DISABLE REPORTS =====================
# Set any format to $false if you don't need it.
# =====================================================================
$generateTxt  = $true
$generateCsv  = $true
$generateHtml = $true


# ====================== REPORT PATHS =================================
# By default, all three files will be placed next to the script.
# To override, set a custom path in quotes.
# =====================================================================
$txtReportPath  = Join-Path $PSScriptRoot "ManualMods-Report.txt"
$csvReportPath  = Join-Path $PSScriptRoot "ManualMods-Report.csv"
$htmlReportPath = Join-Path $PSScriptRoot "ManualMods-Report.html"


# =====================================================================
#  No need to change anything below this line
# =====================================================================

# Ensure proper UTF-8 output in console
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Load System.Web for HtmlEncode (needed for HTML report)
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# Validate path
if (-not (Test-Path $modsPath)) {
    Write-Host "ERROR: Folder not found: $modsPath" -ForegroundColor Red
    Write-Host "Check the modsPath variable at the top of the script." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Mod Organizer 2 - installed mods analysis" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Folder: $modsPath" -ForegroundColor Gray
Write-Host ""


# ---------------------------------------------------------------------
#  Helper functions
# ---------------------------------------------------------------------

# Detect whether a .esp file is actually flagged as ESPFE (light master)
# Per TES4 record spec: byte at offset 9 of the header, 0x02 bit = ESL/ESPFE
function Test-IsEspFe {
    param([string]$EspPath)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($EspPath) | Select-Object -First 16
        if ($bytes.Count -lt 10) { return $false }
        # Byte at index 9 (10th byte) - TES4 record flags
        # 0x02 = ESL flag (light master)
        return ([byte]($bytes[9]) -band 0x02) -ne 0
    }
    catch {
        return $false
    }
}

# Build a Nexus URL from a mod ID
function Get-NexusUrl {
    param([string]$ModId, [string]$Game)
    if ([string]::IsNullOrWhiteSpace($ModId) -or $ModId -eq "0" -or $ModId -eq "-1") {
        return ""
    }
    return "https://www.nexusmods.com/$Game/mods/$ModId"
}

# Extract the mod name from the archive filename (the part before -<modid>-)
# Example: "Paired Animation Improvements-99621-1-0-2-1706671876.7z" -> "Paired Animation Improvements"
function Get-ModNameFromArchive {
    param([string]$ArchiveName, [string]$ModId)
    if ([string]::IsNullOrWhiteSpace($ArchiveName)) { return "" }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($ArchiveName)
    if (-not [string]::IsNullOrWhiteSpace($ModId) -and $ModId -ne "0") {
        # Try to cut everything from -<modid>- onwards
        $pattern = "-$ModId-"
        $idx = $name.IndexOf($pattern)
        if ($idx -gt 0) {
            return $name.Substring(0, $idx)
        }
    }
    # Fallback - cut from the first digit group after a dash
    if ($name -match '^(.+?)-\d+') {
        return $matches[1]
    }
    return $name
}

# Format byte count into human-readable size
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -lt 1KB)  { return "$Bytes B" }
    if ($Bytes -lt 1MB)  { return "{0:N1} KB" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB)  { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N2} GB" -f ($Bytes / 1GB)
}


# ---------------------------------------------------------------------
#  Main analysis
# ---------------------------------------------------------------------

$allMods = Get-ChildItem $modsPath -Directory
$total   = $allMods.Count
$results = New-Object System.Collections.ArrayList

$counter = 0
foreach ($mod in $allMods) {
    $counter++

    # Progress bar
    $percent = [int](($counter / $total) * 100)
    Write-Progress -Activity "Analyzing mods" `
        -Status "$counter / $total - $($mod.Name)" `
        -PercentComplete $percent

    # --- Read meta.ini ---
    $metaFile = Join-Path $mod.FullName "meta.ini"
    $instFile = ""
    $modId    = ""
    $version  = ""
    $gameName = ""
    $hasMeta  = Test-Path $metaFile

    if ($hasMeta) {
        $content = Get-Content $metaFile -Raw -ErrorAction SilentlyContinue
        if ($content -match 'installationFile\s*=\s*(.*)')                  { $instFile = $matches[1].Trim() }
        if ($content -match '(?m)^\s*modid\s*=\s*(.*)')                     { $modId    = $matches[1].Trim() }
        if ($content -match '(?m)^\s*version\s*=\s*(.*)')                   { $version  = $matches[1].Trim() }
        if ($content -match '(?m)^\s*gameName\s*=\s*(.*)')                  { $gameName = $matches[1].Trim() }
    }

    # --- Classify by installation type ---
    $installType = ""
    $reason      = ""
    if (-not $hasMeta) {
        $installType = "NoMeta"
        $reason      = "meta.ini is missing"
    }
    elseif ([string]::IsNullOrWhiteSpace($instFile)) {
        $installType = "NoMeta"
        $reason      = "installationFile is empty or missing"
    }
    elseif ($instFile -match '[\\/]') {
        $installType = "Manual"
        $reason      = "installationFile contains a path to an external archive"
    }
    else {
        $installType = "MO2"
        $reason      = "installationFile is just a filename (installed via MO2)"
    }

    # --- Analyze files inside the mod folder ---
    $espFiles  = @()
    $esmFiles  = @()
    $eslFiles  = @()
    $espFeFiles = @()
    $bsaFiles  = @()
    $ba2Files  = @()
    $bsaSize   = 0L

    try {
        $allFiles = Get-ChildItem $mod.FullName -Recurse -File -ErrorAction SilentlyContinue
        foreach ($f in $allFiles) {
            switch -Wildcard ($f.Extension.ToLower()) {
                ".esp" {
                    $espFiles += $f.Name
                    if (Test-IsEspFe $f.FullName) { $espFeFiles += $f.Name }
                }
                ".esm" { $esmFiles += $f.Name }
                ".esl" { $eslFiles += $f.Name }
                ".bsa" { $bsaFiles += $f.Name; $bsaSize += $f.Length }
                ".ba2" { $ba2Files += $f.Name; $bsaSize += $f.Length }
            }
        }
    }
    catch {
        # Ignore inaccessible files
    }

    $totalPlugins = $espFiles.Count + $esmFiles.Count + $eslFiles.Count
    $hasPlugin    = $totalPlugins -gt 0

    # --- Helper: derived Nexus name and URL ---
    $nexusName = ""
    if ($installType -eq "Manual") {
        # Take only the filename from the full path
        $archiveOnly = Split-Path $instFile -Leaf
        $nexusName   = Get-ModNameFromArchive -ArchiveName $archiveOnly -ModId $modId
    }
    elseif ($installType -eq "MO2" -and -not [string]::IsNullOrWhiteSpace($instFile)) {
        $nexusName = Get-ModNameFromArchive -ArchiveName $instFile -ModId $modId
    }

    $nexusUrl = Get-NexusUrl -ModId $modId -Game $nexusGameSlug

    # --- Build result object ---
    [void]$results.Add([PSCustomObject]@{
        Mod              = $mod.Name
        InstallType      = $installType
        Reason           = $reason
        InstallationFile = $instFile
        ModId            = $modId
        Version          = $version
        GameName         = $gameName
        NexusName        = $nexusName
        NexusUrl         = $nexusUrl
        HasPlugin        = $hasPlugin
        EspCount         = $espFiles.Count
        EsmCount         = $esmFiles.Count
        EslCount         = $eslFiles.Count
        EspFeCount       = $espFeFiles.Count
        Plugins          = ($espFiles + $esmFiles + $eslFiles) -join '; '
        BsaCount         = $bsaFiles.Count + $ba2Files.Count
        BsaSize          = $bsaSize
        BsaSizeFormatted = (Format-Size $bsaSize)
        Bsas             = ($bsaFiles + $ba2Files) -join '; '
    })
}

Write-Progress -Activity "Analyzing mods" -Completed


# ---------------------------------------------------------------------
#  Statistics
# ---------------------------------------------------------------------

$mo2Count    = ($results | Where-Object { $_.InstallType -eq "MO2"    }).Count
$manualCount = ($results | Where-Object { $_.InstallType -eq "Manual" }).Count
$noMetaCount = ($results | Where-Object { $_.InstallType -eq "NoMeta" }).Count

# Suspicious: Manual OR NoMeta - AND no plugins AND no BSA at the same time
# (likely empty folders or broken extraction)
$suspicious = $results | Where-Object {
    ($_.InstallType -eq "Manual" -or $_.InstallType -eq "NoMeta") -and
    -not $_.HasPlugin -and $_.BsaCount -eq 0
}


# ---------------------------------------------------------------------
#  Colored console output
# ---------------------------------------------------------------------

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  STATISTICS" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ("  Total mod folders:                       {0}" -f $total)
Write-Host ("  Installed via MO2 (built-in downloader): {0}" -f $mo2Count)    -ForegroundColor Green
Write-Host ("  Installed manually (external archive):   {0}" -f $manualCount) -ForegroundColor Yellow
Write-Host ("  No meta.ini or no installationFile:      {0}" -f $noMetaCount) -ForegroundColor Yellow
Write-Host ("  SUSPICIOUS (no plugins and no BSA):      {0}" -f $suspicious.Count) -ForegroundColor Red
Write-Host ""

if ($suspicious.Count -gt 0) {
    Write-Host "SUSPICIOUS MODS (likely broken manual install):" -ForegroundColor Red
    foreach ($s in ($suspicious | Sort-Object Mod)) {
        Write-Host ("  - {0}" -f $s.Mod) -ForegroundColor Red
    }
    Write-Host ""
}


# ---------------------------------------------------------------------
#  TXT report
# ---------------------------------------------------------------------

if ($generateTxt) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  MOD ORGANIZER 2 - INSTALLED MODS REPORT")
    [void]$sb.AppendLine("  Date:   $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("  Folder: $modsPath")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("STATISTICS:")
    [void]$sb.AppendLine("  Total mod folders:                       $total")
    [void]$sb.AppendLine("  Installed via MO2 (built-in downloader): $mo2Count")
    [void]$sb.AppendLine("  Installed manually (external archive):   $manualCount")
    [void]$sb.AppendLine("  No meta.ini or no installationFile:      $noMetaCount")
    [void]$sb.AppendLine("  SUSPICIOUS (no plugins and no BSA):      $($suspicious.Count)")
    [void]$sb.AppendLine("")

    # --- Suspicious ---
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  SUSPICIOUS MODS")
    [void]$sb.AppendLine("  (Manual / NoMeta AND no plugins AND no BSA - likely broken)")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    if ($suspicious.Count -eq 0) {
        [void]$sb.AppendLine("  (none)")
    } else {
        foreach ($item in ($suspicious | Sort-Object Mod)) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Install type: $($item.InstallType)")
            [void]$sb.AppendLine("    Reason:       $($item.Reason)")
            if ($item.NexusUrl) { [void]$sb.AppendLine("    Nexus:        $($item.NexusUrl)") }
            if ($item.InstallationFile) { [void]$sb.AppendLine("    Archive:      $($item.InstallationFile)") }
            [void]$sb.AppendLine("")
        }
    }

    # --- List 1: manually installed ---
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  LIST 1: MODS INSTALLED FROM EXTERNAL ARCHIVES")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    $manuals = $results | Where-Object { $_.InstallType -eq "Manual" } | Sort-Object Mod
    if ($manuals.Count -eq 0) {
        [void]$sb.AppendLine("  (none)")
    } else {
        foreach ($item in $manuals) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Archive:  $($item.InstallationFile)")
            if ($item.ModId)    { [void]$sb.AppendLine("    Nexus ID: $($item.ModId)") }
            if ($item.Version)  { [void]$sb.AppendLine("    Version:  $($item.Version)") }
            if ($item.NexusUrl) { [void]$sb.AppendLine("    URL:      $($item.NexusUrl)") }
            $pluginInfo = "    Plugins:  "
            if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount   -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount   -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount   -gt 0) { $parts += "$($item.EslCount) ESL" }
                if ($item.EspFeCount -gt 0) { $parts += "($($item.EspFeCount) of which are ESPFE)" }
                $pluginInfo += ($parts -join ", ")
            } else {
                $pluginInfo += "NONE"
            }
            [void]$sb.AppendLine($pluginInfo)
            if ($item.BsaCount -gt 0) {
                [void]$sb.AppendLine("    BSA/BA2:  $($item.BsaCount) file(s) ($($item.BsaSizeFormatted))")
            }
            [void]$sb.AppendLine("")
        }
    }

    # --- List 2: no meta ---
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  LIST 2: MODS WITHOUT META.INI OR INSTALLATIONFILE")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    $noMetas = $results | Where-Object { $_.InstallType -eq "NoMeta" } | Sort-Object Mod
    if ($noMetas.Count -eq 0) {
        [void]$sb.AppendLine("  (none)")
    } else {
        foreach ($item in $noMetas) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Reason:   $($item.Reason)")
            $pluginInfo = "    Plugins:  "
            if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount -gt 0) { $parts += "$($item.EslCount) ESL" }
                $pluginInfo += ($parts -join ", ")
            } else {
                $pluginInfo += "NONE"
            }
            [void]$sb.AppendLine($pluginInfo)
            if ($item.BsaCount -gt 0) {
                [void]$sb.AppendLine("    BSA/BA2:  $($item.BsaCount) file(s) ($($item.BsaSizeFormatted))")
            }
            [void]$sb.AppendLine("")
        }
    }

    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($txtReportPath, $sb.ToString(), $utf8Bom)
    Write-Host "TXT report:  $txtReportPath" -ForegroundColor Green
}


# ---------------------------------------------------------------------
#  CSV report
# ---------------------------------------------------------------------

if ($generateCsv) {
    # Build objects for CSV (lean schema for easy filtering)
    $csvData = $results | Select-Object Mod, InstallType, ModId, Version, NexusName, NexusUrl,
        HasPlugin, EspCount, EsmCount, EslCount, EspFeCount, BsaCount, BsaSizeFormatted,
        Plugins, Bsas, InstallationFile, Reason | Sort-Object InstallType, Mod

    $csvData | Export-Csv -Path $csvReportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "CSV report:  $csvReportPath" -ForegroundColor Green
}


# ---------------------------------------------------------------------
#  HTML report
# ---------------------------------------------------------------------

if ($generateHtml) {
    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine("<!DOCTYPE html>")
    [void]$html.AppendLine("<html lang='en'><head><meta charset='UTF-8'>")
    [void]$html.AppendLine("<title>MO2 Mods Report</title>")
    [void]$html.AppendLine("<style>")
    [void]$html.AppendLine("body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #1e1e1e; color: #d4d4d4; padding: 20px; }")
    [void]$html.AppendLine("h1 { color: #569cd6; }")
    [void]$html.AppendLine("h2 { color: #4ec9b0; border-bottom: 2px solid #4ec9b0; padding-bottom: 5px; margin-top: 30px; }")
    [void]$html.AppendLine(".stats { background: #252526; padding: 15px; border-radius: 6px; margin-bottom: 20px; }")
    [void]$html.AppendLine(".stats div { margin: 4px 0; }")
    [void]$html.AppendLine(".badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 0.85em; font-weight: bold; }")
    [void]$html.AppendLine(".bg-mo2     { background: #2d7a2d; color: white; }")
    [void]$html.AppendLine(".bg-manual  { background: #b58900; color: white; }")
    [void]$html.AppendLine(".bg-nometa  { background: #cb4b16; color: white; }")
    [void]$html.AppendLine(".bg-danger  { background: #dc322f; color: white; }")
    [void]$html.AppendLine(".bg-ok      { background: #2d7a2d; color: white; }")
    [void]$html.AppendLine("table { border-collapse: collapse; width: 100%; background: #252526; margin-bottom: 30px; }")
    [void]$html.AppendLine("th, td { padding: 8px 12px; text-align: left; border-bottom: 1px solid #3e3e42; vertical-align: top; }")
    [void]$html.AppendLine("th { background: #2d2d30; color: #569cd6; cursor: pointer; user-select: none; }")
    [void]$html.AppendLine("th:hover { background: #3e3e42; }")
    [void]$html.AppendLine("tr:hover { background: #2a2a2d; }")
    [void]$html.AppendLine("a { color: #569cd6; text-decoration: none; }")
    [void]$html.AppendLine("a:hover { text-decoration: underline; }")
    [void]$html.AppendLine(".filter-box { margin-bottom: 10px; }")
    [void]$html.AppendLine(".filter-box input { background: #3c3c3c; color: #d4d4d4; border: 1px solid #555; padding: 6px 10px; border-radius: 4px; width: 300px; }")
    [void]$html.AppendLine(".small { font-size: 0.85em; color: #888; }")
    [void]$html.AppendLine(".danger-row { background: #3a1a1a !important; }")
    [void]$html.AppendLine("</style></head><body>")

    [void]$html.AppendLine("<h1>Mod Organizer 2 - mods report</h1>")
    [void]$html.AppendLine("<div class='small'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>Folder: $modsPath</div>")

    [void]$html.AppendLine("<div class='stats'>")
    [void]$html.AppendLine("<h2 style='margin-top:0; border:none; padding:0;'>Statistics</h2>")
    [void]$html.AppendLine("<div>Total mods: <b>$total</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-mo2'>MO2</span> installed via MO2 (built-in downloader): <b>$mo2Count</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-manual'>Manual</span> installed manually from external archive: <b>$manualCount</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-nometa'>NoMeta</span> no meta.ini / installationFile: <b>$noMetaCount</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-danger'>!</span> Suspicious (no plugins and no BSA): <b>$($suspicious.Count)</b></div>")
    [void]$html.AppendLine("</div>")

    # Helper for a single table
    function Write-HtmlTable {
        param(
            [System.Text.StringBuilder]$Sb,
            [string]$Title,
            [array]$Items,
            [string]$TableId,
            [bool]$HighlightDanger = $false
        )
        [void]$Sb.AppendLine("<h2>$Title (total: $($Items.Count))</h2>")
        if ($Items.Count -eq 0) {
            [void]$Sb.AppendLine("<p class='small'>(none)</p>")
            return
        }
        [void]$Sb.AppendLine("<div class='filter-box'><input type='text' placeholder='Filter by name...' onkeyup=""filterTable('$TableId', this.value)""></div>")
        [void]$Sb.AppendLine("<table id='$TableId'>")
        [void]$Sb.AppendLine("<thead><tr>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 0)"">Mod</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 1)"">Type</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 2)"">Nexus</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 3)"">Version</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 4)"">Plugins</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 5)"">BSA/BA2</th>")
        [void]$Sb.AppendLine("<th>Archive / Reason</th>")
        [void]$Sb.AppendLine("</tr></thead><tbody>")

        foreach ($item in ($Items | Sort-Object Mod)) {
            $rowClass = ""
            if ($HighlightDanger -and -not $item.HasPlugin -and $item.BsaCount -eq 0) {
                $rowClass = " class='danger-row'"
            }

            $typeBadge = switch ($item.InstallType) {
                "MO2"    { "<span class='badge bg-mo2'>MO2</span>" }
                "Manual" { "<span class='badge bg-manual'>Manual</span>" }
                "NoMeta" { "<span class='badge bg-nometa'>NoMeta</span>" }
            }

            $nexusCell = if ($item.NexusUrl) {
                "<a href='$($item.NexusUrl)' target='_blank'>$($item.ModId)</a>"
            } elseif ($item.ModId) {
                $item.ModId
            } else { "-" }

            $pluginCell = if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount   -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount   -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount   -gt 0) { $parts += "$($item.EslCount) ESL" }
                if ($item.EspFeCount -gt 0) { $parts += "<span class='small'>($($item.EspFeCount) ESPFE)</span>" }
                $parts -join ", "
            } else {
                "<span class='badge bg-danger'>NONE</span>"
            }

            $bsaCell = if ($item.BsaCount -gt 0) {
                "$($item.BsaCount) ($($item.BsaSizeFormatted))"
            } else { "-" }

            $reasonCell = if ($item.InstallType -eq "Manual") {
                "<span class='small'>$([System.Web.HttpUtility]::HtmlEncode($item.InstallationFile))</span>"
            } else {
                "<span class='small'>$([System.Web.HttpUtility]::HtmlEncode($item.Reason))</span>"
            }

            $modName = [System.Web.HttpUtility]::HtmlEncode($item.Mod)
            $version = [System.Web.HttpUtility]::HtmlEncode($item.Version)

            [void]$Sb.AppendLine("<tr$rowClass>")
            [void]$Sb.AppendLine("<td>$modName</td>")
            [void]$Sb.AppendLine("<td>$typeBadge</td>")
            [void]$Sb.AppendLine("<td>$nexusCell</td>")
            [void]$Sb.AppendLine("<td>$version</td>")
            [void]$Sb.AppendLine("<td>$pluginCell</td>")
            [void]$Sb.AppendLine("<td>$bsaCell</td>")
            [void]$Sb.AppendLine("<td>$reasonCell</td>")
            [void]$Sb.AppendLine("</tr>")
        }
        [void]$Sb.AppendLine("</tbody></table>")
    }

    Write-HtmlTable -Sb $html -Title "Suspicious mods" `
        -Items $suspicious -TableId "tbl-susp" -HighlightDanger $true

    Write-HtmlTable -Sb $html -Title "Installed manually (external archive)" `
        -Items ($results | Where-Object { $_.InstallType -eq "Manual" }) -TableId "tbl-manual" -HighlightDanger $true

    Write-HtmlTable -Sb $html -Title "No meta.ini or no installationFile" `
        -Items ($results | Where-Object { $_.InstallType -eq "NoMeta" }) -TableId "tbl-nometa" -HighlightDanger $true

    Write-HtmlTable -Sb $html -Title "All mods (including MO2-installed)" `
        -Items $results -TableId "tbl-all" -HighlightDanger $false

    # JS for sorting and filtering
    [void]$html.AppendLine(@"
<script>
function sortTable(tableId, col) {
    var table = document.getElementById(tableId);
    var tbody = table.tBodies[0];
    var rows = Array.from(tbody.rows);
    var asc = table.getAttribute('data-sort-col') != col || table.getAttribute('data-sort-dir') == 'desc';
    rows.sort(function(a, b) {
        var x = a.cells[col].innerText.trim().toLowerCase();
        var y = b.cells[col].innerText.trim().toLowerCase();
        var nx = parseFloat(x), ny = parseFloat(y);
        if (!isNaN(nx) && !isNaN(ny)) { return asc ? nx - ny : ny - nx; }
        return asc ? x.localeCompare(y) : y.localeCompare(x);
    });
    rows.forEach(function(r) { tbody.appendChild(r); });
    table.setAttribute('data-sort-col', col);
    table.setAttribute('data-sort-dir', asc ? 'asc' : 'desc');
}
function filterTable(tableId, query) {
    var q = query.toLowerCase();
    var rows = document.getElementById(tableId).tBodies[0].rows;
    for (var i = 0; i < rows.length; i++) {
        rows[i].style.display = rows[i].cells[0].innerText.toLowerCase().includes(q) ? '' : 'none';
    }
}
</script>
"@)

    [void]$html.AppendLine("</body></html>")

    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($htmlReportPath, $html.ToString(), $utf8Bom)
    Write-Host "HTML report: $htmlReportPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Write-Host ""
