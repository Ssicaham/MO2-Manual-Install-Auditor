# =====================================================================
#  Find-ManualMods.ps1  (версия 2.0)
# ---------------------------------------------------------------------
#  Скрипт анализирует все папки модов в Mod Organizer 2 и
#  классифицирует их по полю installationFile в файле meta.ini,
#  а также собирает дополнительные данные:
#    - наличие плагинов (.esp / .esm / .esl) и их количество
#    - наличие архивов BSA / BA2 и их суммарный размер
#    - Nexus ID мода, версию, имя Nexus-страницы (из имени архива)
#    - готовую ссылку на страницу Nexus
#
#  Генерирует три отчёта:
#    1. ManualMods-Report.txt  - читаемый текстовый отчёт
#    2. ManualMods-Report.csv  - таблица для Excel / фильтрации
#    3. ManualMods-Report.html - HTML с кликабельными ссылками
#
#  Запуск:
#    1. Открой PowerShell в папке со скриптом.
#    2. Если нужно: Unblock-File .\Find-ManualMods.ps1
#    3. Запусти: .\Find-ManualMods.ps1
# =====================================================================


# ====================== ИЗМЕНИ ЭТУ СТРОКУ ============================
# Путь к папке "mods" твоего инстанса MO2.
# Примеры:
#   "D:\MO2\Skyrim Special Edition\mods"
#   "D:\Skyrimmods\mods"
# =====================================================================
$modsPath = "D:\Skyrimmods\mods"


# ====================== ИГРА ДЛЯ ССЫЛОК NEXUS ========================
# Slug игры для построения ссылок вида:
#   https://www.nexusmods.com/<gameSlug>/mods/<id>
# Допустимые значения: skyrimspecialedition, skyrim, fallout4,
# falloutnewvegas, oblivion, morrowind, fallout3, starfield, и т.д.
# =====================================================================
$nexusGameSlug = "skyrimspecialedition"


# ====================== ВКЛЮЧЕНИЕ / ОТКЛЮЧЕНИЕ ОТЧЁТОВ ===============
# Если какой-то формат не нужен - поставь $false
# =====================================================================
$generateTxt  = $true
$generateCsv  = $true
$generateHtml = $true


# ====================== ПУТИ К ОТЧЁТАМ ===============================
# По умолчанию все три файла лягут рядом со скриптом.
# Если надо - переопредели путь вручную в кавычках.
# =====================================================================
$txtReportPath  = Join-Path $PSScriptRoot "ManualMods-Report.txt"
$csvReportPath  = Join-Path $PSScriptRoot "ManualMods-Report.csv"
$htmlReportPath = Join-Path $PSScriptRoot "ManualMods-Report.html"


# =====================================================================
#  Дальше менять ничего не нужно
# =====================================================================

# Корректный вывод кириллицы в консоль
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Загружаем System.Web для HtmlEncode (нужно для HTML-отчёта)
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# Проверка пути
if (-not (Test-Path $modsPath)) {
    Write-Host "ОШИБКА: Папка не найдена: $modsPath" -ForegroundColor Red
    Write-Host "Проверь переменную modsPath в начале скрипта." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Mod Organizer 2 - анализ установленных модов" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Папка: $modsPath" -ForegroundColor Gray
Write-Host ""


# ---------------------------------------------------------------------
#  Вспомогательные функции
# ---------------------------------------------------------------------

# Определить, является ли .esp на самом деле ESPFE (light master flag)
# По спецификации формата TES4: 10-й байт заголовка = 0x02 -> ESL/ESPFE
function Test-IsEspFe {
    param([string]$EspPath)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($EspPath) | Select-Object -First 16
        if ($bytes.Count -lt 10) { return $false }
        # Байт по индексу 9 (10-й по счёту) - флаги записи TES4
        # 0x02 = ESL flag (light master)
        return ([byte]($bytes[9]) -band 0x02) -ne 0
    }
    catch {
        return $false
    }
}

# Построить ссылку на Nexus из ID
function Get-NexusUrl {
    param([string]$ModId, [string]$Game)
    if ([string]::IsNullOrWhiteSpace($ModId) -or $ModId -eq "0" -or $ModId -eq "-1") {
        return ""
    }
    return "https://www.nexusmods.com/$Game/mods/$ModId"
}

# Извлечь имя мода из имени архива (то, что до -<modid>-)
# Пример: "Paired Animation Improvements-99621-1-0-2-1706671876.7z" -> "Paired Animation Improvements"
function Get-ModNameFromArchive {
    param([string]$ArchiveName, [string]$ModId)
    if ([string]::IsNullOrWhiteSpace($ArchiveName)) { return "" }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($ArchiveName)
    if (-not [string]::IsNullOrWhiteSpace($ModId) -and $ModId -ne "0") {
        # Пробуем отрезать всё начиная с -<modid>-
        $pattern = "-$ModId-"
        $idx = $name.IndexOf($pattern)
        if ($idx -gt 0) {
            return $name.Substring(0, $idx)
        }
    }
    # Запасной вариант - отрезать всё начиная с первой группы цифр после дефиса
    if ($name -match '^(.+?)-\d+') {
        return $matches[1]
    }
    return $name
}

# Форматирование размера в человекочитаемый вид
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -lt 1KB)  { return "$Bytes B" }
    if ($Bytes -lt 1MB)  { return "{0:N1} KB" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB)  { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N2} GB" -f ($Bytes / 1GB)
}


# ---------------------------------------------------------------------
#  Основной анализ
# ---------------------------------------------------------------------

$allMods = Get-ChildItem $modsPath -Directory
$total   = $allMods.Count
$results = New-Object System.Collections.ArrayList

$counter = 0
foreach ($mod in $allMods) {
    $counter++

    # Прогресс-бар
    $percent = [int](($counter / $total) * 100)
    Write-Progress -Activity "Анализирую моды" `
        -Status "$counter / $total - $($mod.Name)" `
        -PercentComplete $percent

    # --- Чтение meta.ini ---
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

    # --- Классификация по типу установки ---
    $installType = ""
    $reason      = ""
    if (-not $hasMeta) {
        $installType = "NoMeta"
        $reason      = "meta.ini отсутствует"
    }
    elseif ([string]::IsNullOrWhiteSpace($instFile)) {
        $installType = "NoMeta"
        $reason      = "installationFile пустое или отсутствует"
    }
    elseif ($instFile -match '[\\/]') {
        $installType = "Manual"
        $reason      = "installationFile содержит путь к внешнему архиву"
    }
    else {
        $installType = "MO2"
        $reason      = "installationFile - только имя файла (штатная установка)"
    }

    # --- Анализ файлов внутри мода ---
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
        # Игнорируем недоступные файлы
    }

    $totalPlugins = $espFiles.Count + $esmFiles.Count + $eslFiles.Count
    $hasPlugin    = $totalPlugins -gt 0

    # --- Вспомогательное: имя по Nexus и URL ---
    $nexusName = ""
    if ($installType -eq "Manual") {
        # Берём только имя файла из полного пути
        $archiveOnly = Split-Path $instFile -Leaf
        $nexusName   = Get-ModNameFromArchive -ArchiveName $archiveOnly -ModId $modId
    }
    elseif ($installType -eq "MO2" -and -not [string]::IsNullOrWhiteSpace($instFile)) {
        $nexusName = Get-ModNameFromArchive -ArchiveName $instFile -ModId $modId
    }

    $nexusUrl = Get-NexusUrl -ModId $modId -Game $nexusGameSlug

    # --- Сборка объекта ---
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

Write-Progress -Activity "Анализирую моды" -Completed


# ---------------------------------------------------------------------
#  Подсчёт статистики
# ---------------------------------------------------------------------

$mo2Count    = ($results | Where-Object { $_.InstallType -eq "MO2"    }).Count
$manualCount = ($results | Where-Object { $_.InstallType -eq "Manual" }).Count
$noMetaCount = ($results | Where-Object { $_.InstallType -eq "NoMeta" }).Count

# Подозрительные: ручные ИЛИ без meta - и без плагинов и без BSA одновременно
# (чисто пустые папки или явные косяки распаковки)
$suspicious = $results | Where-Object {
    ($_.InstallType -eq "Manual" -or $_.InstallType -eq "NoMeta") -and
    -not $_.HasPlugin -and $_.BsaCount -eq 0
}


# ---------------------------------------------------------------------
#  Цветной вывод в консоль
# ---------------------------------------------------------------------

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  СТАТИСТИКА" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ("  Всего папок модов:                       {0}" -f $total)
Write-Host ("  Установлены штатно через MO2:            {0}" -f $mo2Count)    -ForegroundColor Green
Write-Host ("  Установлены вручную (путь к архиву):     {0}" -f $manualCount) -ForegroundColor Yellow
Write-Host ("  Без meta.ini или без installationFile:   {0}" -f $noMetaCount) -ForegroundColor Yellow
Write-Host ("  ПОДОЗРИТЕЛЬНЫЕ (без плагинов и BSA):     {0}" -f $suspicious.Count) -ForegroundColor Red
Write-Host ""

if ($suspicious.Count -gt 0) {
    Write-Host "ПОДОЗРИТЕЛЬНЫЕ МОДЫ (вероятно, битая ручная установка):" -ForegroundColor Red
    foreach ($s in ($suspicious | Sort-Object Mod)) {
        Write-Host ("  - {0}" -f $s.Mod) -ForegroundColor Red
    }
    Write-Host ""
}


# ---------------------------------------------------------------------
#  TXT отчёт
# ---------------------------------------------------------------------

if ($generateTxt) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  ОТЧЁТ ПО МОДАМ MOD ORGANIZER 2")
    [void]$sb.AppendLine("  Дата:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("  Папка: $modsPath")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("СТАТИСТИКА:")
    [void]$sb.AppendLine("  Всего папок модов:                       $total")
    [void]$sb.AppendLine("  Установлены штатно через MO2:            $mo2Count")
    [void]$sb.AppendLine("  Установлены вручную (путь к архиву):     $manualCount")
    [void]$sb.AppendLine("  Без meta.ini или без installationFile:   $noMetaCount")
    [void]$sb.AppendLine("  ПОДОЗРИТЕЛЬНЫЕ (без плагинов и BSA):     $($suspicious.Count)")
    [void]$sb.AppendLine("")

    # --- Подозрительные ---
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  ПОДОЗРИТЕЛЬНЫЕ МОДЫ")
    [void]$sb.AppendLine("  (ручные / без meta И без плагинов И без BSA - возможный косяк)")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    if ($suspicious.Count -eq 0) {
        [void]$sb.AppendLine("  (нет таких модов)")
    } else {
        foreach ($item in ($suspicious | Sort-Object Mod)) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Тип установки: $($item.InstallType)")
            [void]$sb.AppendLine("    Причина:       $($item.Reason)")
            if ($item.NexusUrl) { [void]$sb.AppendLine("    Nexus:         $($item.NexusUrl)") }
            if ($item.InstallationFile) { [void]$sb.AppendLine("    Архив:         $($item.InstallationFile)") }
            [void]$sb.AppendLine("")
        }
    }

    # --- Список 1: ручные ---
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  СПИСОК 1: МОДЫ, УСТАНОВЛЕННЫЕ ИЗ ВНЕШНИХ АРХИВОВ")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    $manuals = $results | Where-Object { $_.InstallType -eq "Manual" } | Sort-Object Mod
    if ($manuals.Count -eq 0) {
        [void]$sb.AppendLine("  (нет таких модов)")
    } else {
        foreach ($item in $manuals) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Архив:    $($item.InstallationFile)")
            if ($item.ModId)    { [void]$sb.AppendLine("    Nexus ID: $($item.ModId)") }
            if ($item.Version)  { [void]$sb.AppendLine("    Версия:   $($item.Version)") }
            if ($item.NexusUrl) { [void]$sb.AppendLine("    URL:      $($item.NexusUrl)") }
            $pluginInfo = "    Плагины:  "
            if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount   -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount   -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount   -gt 0) { $parts += "$($item.EslCount) ESL" }
                if ($item.EspFeCount -gt 0) { $parts += "(из них $($item.EspFeCount) ESPFE)" }
                $pluginInfo += ($parts -join ", ")
            } else {
                $pluginInfo += "НЕТ"
            }
            [void]$sb.AppendLine($pluginInfo)
            if ($item.BsaCount -gt 0) {
                [void]$sb.AppendLine("    BSA/BA2:  $($item.BsaCount) шт. ($($item.BsaSizeFormatted))")
            }
            [void]$sb.AppendLine("")
        }
    }

    # --- Список 2: без meta ---
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  СПИСОК 2: МОДЫ БЕЗ META.INI ИЛИ БЕЗ INSTALLATIONFILE")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    $noMetas = $results | Where-Object { $_.InstallType -eq "NoMeta" } | Sort-Object Mod
    if ($noMetas.Count -eq 0) {
        [void]$sb.AppendLine("  (нет таких модов)")
    } else {
        foreach ($item in $noMetas) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Причина:  $($item.Reason)")
            $pluginInfo = "    Плагины:  "
            if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount -gt 0) { $parts += "$($item.EslCount) ESL" }
                $pluginInfo += ($parts -join ", ")
            } else {
                $pluginInfo += "НЕТ"
            }
            [void]$sb.AppendLine($pluginInfo)
            if ($item.BsaCount -gt 0) {
                [void]$sb.AppendLine("    BSA/BA2:  $($item.BsaCount) шт. ($($item.BsaSizeFormatted))")
            }
            [void]$sb.AppendLine("")
        }
    }

    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($txtReportPath, $sb.ToString(), $utf8Bom)
    Write-Host "TXT отчёт:  $txtReportPath" -ForegroundColor Green
}


# ---------------------------------------------------------------------
#  CSV отчёт
# ---------------------------------------------------------------------

if ($generateCsv) {
    # Готовим объекты для CSV (без длинных списков, чтобы было удобно фильтровать)
    $csvData = $results | Select-Object Mod, InstallType, ModId, Version, NexusName, NexusUrl,
        HasPlugin, EspCount, EsmCount, EslCount, EspFeCount, BsaCount, BsaSizeFormatted,
        Plugins, Bsas, InstallationFile, Reason | Sort-Object InstallType, Mod

    $csvData | Export-Csv -Path $csvReportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "CSV отчёт:  $csvReportPath" -ForegroundColor Green
}


# ---------------------------------------------------------------------
#  HTML отчёт
# ---------------------------------------------------------------------

if ($generateHtml) {
    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine("<!DOCTYPE html>")
    [void]$html.AppendLine("<html lang='ru'><head><meta charset='UTF-8'>")
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

    [void]$html.AppendLine("<h1>Mod Organizer 2 - отчёт по модам</h1>")
    [void]$html.AppendLine("<div class='small'>Сгенерировано: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>Папка: $modsPath</div>")

    [void]$html.AppendLine("<div class='stats'>")
    [void]$html.AppendLine("<h2 style='margin-top:0; border:none; padding:0;'>Статистика</h2>")
    [void]$html.AppendLine("<div>Всего модов: <b>$total</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-mo2'>MO2</span> штатно через MO2: <b>$mo2Count</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-manual'>Manual</span> вручную из внешнего архива: <b>$manualCount</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-nometa'>NoMeta</span> без meta.ini / installationFile: <b>$noMetaCount</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-danger'>!</span> Подозрительные (без плагинов и BSA): <b>$($suspicious.Count)</b></div>")
    [void]$html.AppendLine("</div>")

    # Функция для одной таблицы
    function Write-HtmlTable {
        param(
            [System.Text.StringBuilder]$Sb,
            [string]$Title,
            [array]$Items,
            [string]$TableId,
            [bool]$HighlightDanger = $false
        )
        [void]$Sb.AppendLine("<h2>$Title (всего: $($Items.Count))</h2>")
        if ($Items.Count -eq 0) {
            [void]$Sb.AppendLine("<p class='small'>(нет таких модов)</p>")
            return
        }
        [void]$Sb.AppendLine("<div class='filter-box'><input type='text' placeholder='Фильтр по имени...' onkeyup=""filterTable('$TableId', this.value)""></div>")
        [void]$Sb.AppendLine("<table id='$TableId'>")
        [void]$Sb.AppendLine("<thead><tr>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 0)"">Мод</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 1)"">Тип</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 2)"">Nexus</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 3)"">Версия</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 4)"">Плагины</th>")
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId', 5)"">BSA/BA2</th>")
        [void]$Sb.AppendLine("<th>Архив / Причина</th>")
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
                "<span class='badge bg-danger'>НЕТ</span>"
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

    Write-HtmlTable -Sb $html -Title "Подозрительные моды" `
        -Items $suspicious -TableId "tbl-susp" -HighlightDanger $true

    Write-HtmlTable -Sb $html -Title "Установлены вручную (внешний архив)" `
        -Items ($results | Where-Object { $_.InstallType -eq "Manual" }) -TableId "tbl-manual" -HighlightDanger $true

    Write-HtmlTable -Sb $html -Title "Без meta.ini или без installationFile" `
        -Items ($results | Where-Object { $_.InstallType -eq "NoMeta" }) -TableId "tbl-nometa" -HighlightDanger $true

    Write-HtmlTable -Sb $html -Title "Все моды (включая штатные)" `
        -Items $results -TableId "tbl-all" -HighlightDanger $false

    # JS для сортировки и фильтра
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
    Write-Host "HTML отчёт: $htmlReportPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "Готово." -ForegroundColor Cyan
Write-Host ""
