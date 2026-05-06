# =====================================================================
#  MO2 Manual Install Auditor - GUI version (v1.2)
#  Find-ManualMods-GUI.ps1
# ---------------------------------------------------------------------
#  WPF wrapper around the analysis logic from Find-ManualMods.ps1.
#  Run as a PowerShell script, or compile to .exe with ps2exe:
#    Install-Module ps2exe -Scope CurrentUser
#    Invoke-PS2EXE .\Find-ManualMods-GUI.ps1 .\MO2ManualInstallAuditor.exe -noConsole
#
#  v1.2 changes:
#    - Optional MO2 profile folder selection - reads modlist.txt to extract
#      load order (priority) and enabled/disabled state for every mod
#    - Reports (TXT / CSV / HTML) include Priority and Enabled columns
#    - HTML "All mods" table is sorted by priority by default
#  v1.1 changes:
#    - Settings persistence in %APPDATA%\MO2ManualInstallAuditor\settings.json
#    - Dual language support (English / Russian), extensible
# =====================================================================

# UTF-8 console encoding (skipped silently when running as ps2exe -noConsole)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue


# =====================================================================
#  TRANSLATIONS
# ---------------------------------------------------------------------
#  To add a new language:
#    1. Copy any block (e.g. 'en' = @{...}) under a new key like 'de'.
#    2. Translate every value on the right side of '='.
#    3. Add the language to $LanguageList below with its display name.
#    Keys (left side) must NOT be changed.
# =====================================================================

$Translations = @{
    'en' = @{
        'window_title'         = 'MO2 Manual Install Auditor'
        'app_subtitle'         = "Detects manually installed mods in Mod Organizer 2 and generates diagnostic reports."
        'lbl_language'         = 'Language:'
        'lbl_mods_path'        = "Path to MO2 'mods' folder:"
        'lbl_profile_path'     = "Path to MO2 profile folder (optional, for load order):"
        'btn_browse'           = 'Browse...'
        'browse_dialog_desc'   = "Select your MO2 'mods' folder"
        'browse_profile_desc'  = "Select an MO2 profile folder (containing modlist.txt)"
        'lbl_game_slug'        = 'Nexus game slug:'
        'hint_game_slug'       = 'e.g. skyrimspecialedition, fallout4, starfield, oblivion'
        'lbl_reports'          = 'Reports:'
        'btn_run'              = 'Run Analysis'
        'btn_open_folder'      = 'Open Reports Folder'
        'status_ready'         = 'Ready.'
        'status_scanning'      = 'Scanning...'
        'status_analyzing'     = 'Analyzing: {0} / {1}'
        'status_done'          = 'Done. {0} suspicious mod(s) found.'
        'status_failed'        = 'Failed.'
        'status_nothing'       = 'Nothing to scan.'
        'status_error_folder'  = 'Error: folder not found.'
        'msg_missing_path'     = "Please specify the path to the MO2 'mods' folder."
        'msg_missing_path_t'   = 'Missing path'
        'log_scanning'         = 'Scanning: {0}'
        'log_no_folders'       = 'No mod folders found in {0}'
        'log_found_folders'    = 'Found {0} mod folders. Analyzing...'
        'log_error_folder'     = 'ERROR: Folder not found: {0}'
        'log_modlist_found'    = 'modlist.txt found - extracted priority for {0} mod(s).'
        'log_modlist_missing'  = 'modlist.txt not found in profile folder - priority data skipped.'
        'log_no_profile'       = 'No profile folder set - load order data will not be included.'
        'log_stats_header'     = '=== STATISTICS ==='
        'log_stats_total'      = '  Total mods:                         {0}'
        'log_stats_mo2'        = '  Installed via MO2:                  {0}'
        'log_stats_manual'     = '  Installed manually (external path): {0}'
        'log_stats_nometa'     = '  No meta.ini / no installationFile:  {0}'
        'log_stats_susp'       = '  Suspicious (no plugins and no BSA): {0}'
        'log_stats_enabled'    = '  Enabled in profile:                 {0}'
        'log_stats_disabled'   = '  Disabled in profile:                {0}'
        'log_susp_header'      = 'SUSPICIOUS MODS (likely broken manual install):'
        'log_txt_saved'        = 'TXT report saved.'
        'log_csv_saved'        = 'CSV report saved.'
        'log_html_saved'       = 'HTML report saved.'
        'log_reports_to'       = 'Reports saved to: {0}'
        'log_error_generic'    = 'ERROR: {0}'
    }
    'ru' = @{
        'window_title'         = 'MO2 Аудитор ручной установки'
        'app_subtitle'         = 'Находит вручную установленные моды в Mod Organizer 2 и формирует диагностические отчёты.'
        'lbl_language'         = 'Язык:'
        'lbl_mods_path'        = "Путь к папке 'mods' вашего инстанса MO2:"
        'lbl_profile_path'     = 'Путь к папке профиля MO2 (опционально, для порядка загрузки):'
        'btn_browse'           = 'Обзор...'
        'browse_dialog_desc'   = "Выберите папку 'mods' вашего инстанса MO2"
        'browse_profile_desc'  = 'Выберите папку профиля MO2 (содержащую modlist.txt)'
        'lbl_game_slug'        = 'Slug игры на Nexus:'
        'hint_game_slug'       = 'Например: skyrimspecialedition, fallout4, starfield, oblivion'
        'lbl_reports'          = 'Отчёты:'
        'btn_run'              = 'Запустить анализ'
        'btn_open_folder'      = 'Открыть папку с отчётами'
        'status_ready'         = 'Готов к работе.'
        'status_scanning'      = 'Сканирование...'
        'status_analyzing'     = 'Анализ: {0} / {1}'
        'status_done'          = 'Готово. Найдено подозрительных модов: {0}.'
        'status_failed'        = 'Ошибка выполнения.'
        'status_nothing'       = 'Сканировать нечего.'
        'status_error_folder'  = 'Ошибка: папка не найдена.'
        'msg_missing_path'     = "Укажите путь к папке 'mods' вашего инстанса MO2."
        'msg_missing_path_t'   = 'Не указан путь'
        'log_scanning'         = 'Сканирование: {0}'
        'log_no_folders'       = 'Папки модов не найдены в {0}'
        'log_found_folders'    = 'Найдено {0} папок модов. Начинаю анализ...'
        'log_error_folder'     = 'ОШИБКА: Папка не найдена: {0}'
        'log_modlist_found'    = 'Файл modlist.txt найден - получены приоритеты для {0} мод(ов).'
        'log_modlist_missing'  = 'Файл modlist.txt не найден в папке профиля - данные о приоритете пропущены.'
        'log_no_profile'       = 'Папка профиля не указана - данные о порядке загрузки не будут включены.'
        'log_stats_header'     = '=== СТАТИСТИКА ==='
        'log_stats_total'      = '  Всего модов:                            {0}'
        'log_stats_mo2'        = '  Установлены штатно через MO2:           {0}'
        'log_stats_manual'     = '  Установлены вручную (внешний путь):     {0}'
        'log_stats_nometa'     = '  Без meta.ini / без installationFile:    {0}'
        'log_stats_susp'       = '  Подозрительные (без плагинов и BSA):    {0}'
        'log_stats_enabled'    = '  Включено в профиле:                     {0}'
        'log_stats_disabled'   = '  Отключено в профиле:                    {0}'
        'log_susp_header'      = 'ПОДОЗРИТЕЛЬНЫЕ МОДЫ (вероятно, битая ручная установка):'
        'log_txt_saved'        = 'Отчёт TXT сохранён.'
        'log_csv_saved'        = 'Отчёт CSV сохранён.'
        'log_html_saved'       = 'Отчёт HTML сохранён.'
        'log_reports_to'       = 'Отчёты сохранены в: {0}'
        'log_error_generic'    = 'ОШИБКА: {0}'
    }
}

$LanguageList = @(
    [PSCustomObject]@{ Code = 'en'; Display = 'English'  }
    [PSCustomObject]@{ Code = 'ru'; Display = 'Русский'  }
)

$script:CurrentLanguage = 'en'

function Get-Text {
    param([string]$Key)
    $dict = $Translations[$script:CurrentLanguage]
    if (-not $dict) { $dict = $Translations['en'] }
    if ($dict.ContainsKey($Key)) { return $dict[$Key] }
    if ($Translations['en'].ContainsKey($Key)) { return $Translations['en'][$Key] }
    return "[$Key]"
}


# =====================================================================
#  SETTINGS PERSISTENCE
# =====================================================================

$settingsDir  = Join-Path $env:APPDATA "MO2ManualInstallAuditor"
$settingsFile = Join-Path $settingsDir "settings.json"

function Load-Settings {
    $defaults = [PSCustomObject]@{
        modsPath     = ""
        profilePath  = ""
        gameSlug     = "skyrimspecialedition"
        generateTxt  = $true
        generateCsv  = $true
        generateHtml = $true
        language     = "en"
    }
    if (-not (Test-Path $settingsFile)) { return $defaults }
    try {
        $loaded = Get-Content $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $defaults.PSObject.Properties.Name) {
            if ($null -ne $loaded.$prop) { $defaults.$prop = $loaded.$prop }
        }
        return $defaults
    }
    catch { return $defaults }
}

function Save-Settings {
    param($Settings)
    try {
        if (-not (Test-Path $settingsDir)) {
            [void](New-Item -ItemType Directory -Path $settingsDir -Force)
        }
        $json = $Settings | ConvertTo-Json -Depth 4
        $utf8Bom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($settingsFile, $json, $utf8Bom)
    }
    catch { }
}

$script:Settings = Load-Settings
$script:CurrentLanguage = $script:Settings.language
if (-not $Translations.ContainsKey($script:CurrentLanguage)) { $script:CurrentLanguage = 'en' }


# =====================================================================
#  WPF window definition (XAML)
# =====================================================================

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MO2 Manual Install Auditor"
        Width="820" Height="720"
        Background="#1E1E1E" Foreground="#D4D4D4"
        FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Background" Value="#2D2D30"/>
      <Setter Property="Foreground" Value="#D4D4D4"/>
      <Setter Property="BorderBrush" Value="#3E3E42"/>
      <Setter Property="Padding" Value="10,5"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#2D2D30"/>
      <Setter Property="Foreground" Value="#D4D4D4"/>
      <Setter Property="BorderBrush" Value="#3E3E42"/>
      <Setter Property="Padding" Value="4"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#D4D4D4"/>
      <Setter Property="Margin" Value="0,4,12,4"/>
    </Style>
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#9CDCFE"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="#2D2D30"/>
      <Setter Property="Foreground" Value="#000000"/>
      <Setter Property="BorderBrush" Value="#3E3E42"/>
    </Style>
  </Window.Resources>

  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid Grid.Row="0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="TitleText" Grid.Column="0" Text="MO2 Manual Install Auditor"
                 FontSize="20" FontWeight="Bold" Foreground="#569CD6" Margin="0,0,0,4"/>
      <Label x:Name="LangLabel" Grid.Column="1" Content="Language:" VerticalAlignment="Center" Margin="0,0,4,0"/>
      <ComboBox x:Name="LanguageCombo" Grid.Column="2" Width="120" VerticalAlignment="Center"/>
    </Grid>

    <TextBlock x:Name="SubtitleText" Grid.Row="1" Text="" Foreground="#888" Margin="0,0,0,12"/>

    <Label x:Name="ModsPathLabel" Grid.Row="2" Content=""/>
    <Grid Grid.Row="3" Margin="0,0,0,8">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBox x:Name="ModsPathBox" Grid.Column="0"/>
      <Button x:Name="BrowseButton" Grid.Column="1" Content="Browse..." Margin="6,0,0,0" Width="110"/>
    </Grid>

    <Label x:Name="ProfilePathLabel" Grid.Row="4" Content=""/>
    <Grid Grid.Row="5" Margin="0,0,0,8">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBox x:Name="ProfilePathBox" Grid.Column="0"/>
      <Button x:Name="BrowseProfileButton" Grid.Column="1" Content="Browse..." Margin="6,0,0,0" Width="110"/>
    </Grid>

    <Grid Grid.Row="6" Margin="0,0,0,8">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="200"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Label x:Name="GameSlugLabel" Grid.Column="0" Content="" VerticalAlignment="Center"/>
      <TextBox x:Name="GameSlugBox" Grid.Column="1" Text="skyrimspecialedition"/>
      <TextBlock x:Name="GameSlugHint" Grid.Column="2" Foreground="#888" VerticalAlignment="Center" Margin="10,0,0,0" Text=""/>
    </Grid>

    <StackPanel Grid.Row="7" Orientation="Horizontal" Margin="0,4,0,10">
      <Label x:Name="ReportsLabel" Content="" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <CheckBox x:Name="TxtCheck"  Content="TXT"  IsChecked="True"/>
      <CheckBox x:Name="CsvCheck"  Content="CSV"  IsChecked="True"/>
      <CheckBox x:Name="HtmlCheck" Content="HTML" IsChecked="True"/>
      <Button   x:Name="RunButton" Content="" Margin="20,0,0,0" Width="160"/>
      <Button   x:Name="OpenFolderButton" Content="" Margin="8,0,0,0" Width="200" IsEnabled="False"/>
    </StackPanel>

    <ProgressBar x:Name="Progress" Grid.Row="8" Height="18" Margin="0,0,0,8"
                 Background="#2D2D30" Foreground="#569CD6" BorderBrush="#3E3E42"
                 VerticalAlignment="Top"/>

    <TextBox x:Name="LogBox" Grid.Row="8" Margin="0,30,0,0"
             IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
             FontFamily="Consolas" FontSize="12" Background="#161616"/>

    <TextBlock x:Name="StatusLine" Grid.Row="9" Foreground="#888" Margin="0,8,0,0" Text=""/>
  </Grid>
</Window>
"@


# =====================================================================
#  Load XAML and wire up controls
# =====================================================================

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$titleText           = $window.FindName("TitleText")
$subtitleText        = $window.FindName("SubtitleText")
$langLabel           = $window.FindName("LangLabel")
$languageCombo       = $window.FindName("LanguageCombo")
$modsPathLabel       = $window.FindName("ModsPathLabel")
$modsPathBox         = $window.FindName("ModsPathBox")
$browseButton        = $window.FindName("BrowseButton")
$profilePathLabel    = $window.FindName("ProfilePathLabel")
$profilePathBox      = $window.FindName("ProfilePathBox")
$browseProfileButton = $window.FindName("BrowseProfileButton")
$gameSlugLabel       = $window.FindName("GameSlugLabel")
$gameSlugBox         = $window.FindName("GameSlugBox")
$gameSlugHint        = $window.FindName("GameSlugHint")
$reportsLabel        = $window.FindName("ReportsLabel")
$txtCheck            = $window.FindName("TxtCheck")
$csvCheck            = $window.FindName("CsvCheck")
$htmlCheck           = $window.FindName("HtmlCheck")
$runButton           = $window.FindName("RunButton")
$openFolderButton    = $window.FindName("OpenFolderButton")
$progress            = $window.FindName("Progress")
$logBox              = $window.FindName("LogBox")
$statusLine          = $window.FindName("StatusLine")

# Output folder = next to script/exe
$outputDir = $null
if ($PSScriptRoot) { $outputDir = $PSScriptRoot }
if (-not $outputDir) {
    try {
        $entryAsm = [System.Reflection.Assembly]::GetEntryAssembly()
        if ($entryAsm -and $entryAsm.Location) {
            $outputDir = [System.IO.Path]::GetDirectoryName($entryAsm.Location)
        }
    } catch { }
}
if (-not $outputDir) {
    try {
        $procPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($procPath) { $outputDir = [System.IO.Path]::GetDirectoryName($procPath) }
    } catch { }
}
if (-not $outputDir) { $outputDir = [Environment]::CurrentDirectory }


# =====================================================================
#  Apply translations to all UI elements
# =====================================================================

function Apply-Translations {
    $window.Title              = Get-Text 'window_title'
    $titleText.Text            = Get-Text 'window_title'
    $subtitleText.Text         = Get-Text 'app_subtitle'
    $langLabel.Content         = Get-Text 'lbl_language'
    $modsPathLabel.Content     = Get-Text 'lbl_mods_path'
    $profilePathLabel.Content  = Get-Text 'lbl_profile_path'
    $browseButton.Content      = Get-Text 'btn_browse'
    $browseProfileButton.Content = Get-Text 'btn_browse'
    $gameSlugLabel.Content     = Get-Text 'lbl_game_slug'
    $gameSlugHint.Text         = Get-Text 'hint_game_slug'
    $reportsLabel.Content      = Get-Text 'lbl_reports'
    $runButton.Content         = Get-Text 'btn_run'
    $openFolderButton.Content  = Get-Text 'btn_open_folder'
    if (-not $statusLine.Text -or $statusLine.Tag -eq 'ready') {
        $statusLine.Text = Get-Text 'status_ready'
        $statusLine.Tag  = 'ready'
    }
}


# =====================================================================
#  Initialize controls from settings
# =====================================================================

foreach ($lang in $LanguageList) {
    [void]$languageCombo.Items.Add($lang.Display)
}
$idx = 0
for ($i = 0; $i -lt $LanguageList.Count; $i++) {
    if ($LanguageList[$i].Code -eq $script:CurrentLanguage) { $idx = $i; break }
}
$languageCombo.SelectedIndex = $idx

$modsPathBox.Text     = $script:Settings.modsPath
$profilePathBox.Text  = $script:Settings.profilePath
$gameSlugBox.Text     = $script:Settings.gameSlug
$txtCheck.IsChecked   = [bool]$script:Settings.generateTxt
$csvCheck.IsChecked   = [bool]$script:Settings.generateCsv
$htmlCheck.IsChecked  = [bool]$script:Settings.generateHtml

Apply-Translations


# =====================================================================
#  UI helpers
# =====================================================================

function Write-Log {
    param([string]$Text, [string]$Color = "#D4D4D4")
    $logBox.Dispatcher.Invoke([Action]{
        $logBox.AppendText("$Text`r`n")
        $logBox.ScrollToEnd()
    })
}

function Set-Status {
    param([string]$Text, [string]$Tag = "")
    $statusLine.Dispatcher.Invoke([Action]{
        $statusLine.Text = $Text
        $statusLine.Tag = $Tag
    })
}

function Set-Progress {
    param([int]$Value)
    $progress.Dispatcher.Invoke([Action]{ $progress.Value = $Value })
}


# =====================================================================
#  Settings save
# =====================================================================

function Persist-Settings {
    $script:Settings = [PSCustomObject]@{
        modsPath     = $modsPathBox.Text.Trim()
        profilePath  = $profilePathBox.Text.Trim()
        gameSlug     = $gameSlugBox.Text.Trim()
        generateTxt  = [bool]$txtCheck.IsChecked
        generateCsv  = [bool]$csvCheck.IsChecked
        generateHtml = [bool]$htmlCheck.IsChecked
        language     = $script:CurrentLanguage
    }
    Save-Settings $script:Settings
}


# =====================================================================
#  Event handlers
# =====================================================================

$languageCombo.Add_SelectionChanged({
    $sel = $languageCombo.SelectedIndex
    if ($sel -ge 0 -and $sel -lt $LanguageList.Count) {
        $script:CurrentLanguage = $LanguageList[$sel].Code
        Apply-Translations
        Persist-Settings
    }
})

$modsPathBox.Add_LostFocus({ Persist-Settings })
$profilePathBox.Add_LostFocus({ Persist-Settings })
$gameSlugBox.Add_LostFocus({ Persist-Settings })
$txtCheck.Add_Click({ Persist-Settings })
$csvCheck.Add_Click({ Persist-Settings })
$htmlCheck.Add_Click({ Persist-Settings })

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = Get-Text 'browse_dialog_desc'
    $dialog.ShowNewFolderButton = $false
    if (-not [string]::IsNullOrWhiteSpace($modsPathBox.Text)) {
        if (Test-Path $modsPathBox.Text) { $dialog.SelectedPath = $modsPathBox.Text }
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $modsPathBox.Text = $dialog.SelectedPath
        Persist-Settings
    }
})

$browseProfileButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = Get-Text 'browse_profile_desc'
    $dialog.ShowNewFolderButton = $false
    if (-not [string]::IsNullOrWhiteSpace($profilePathBox.Text)) {
        if (Test-Path $profilePathBox.Text) { $dialog.SelectedPath = $profilePathBox.Text }
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $profilePathBox.Text = $dialog.SelectedPath
        Persist-Settings
    }
})

$openFolderButton.Add_Click({
    if (Test-Path $outputDir) { Start-Process explorer.exe $outputDir }
})


# =====================================================================
#  Helper: parse modlist.txt into priority/enabled hashtable
# ---------------------------------------------------------------------
#  modlist.txt format (lines from top to bottom):
#    +ModName  -> enabled
#    -ModName  -> disabled
#    *Lines or # comments -> ignored
#  In the file the LAST line has the HIGHEST load priority (loads last,
#  overrides earlier mods). To match how MO2 displays the list (top of UI
#  = highest priority), we reverse the order so that priority 1 = top.
# =====================================================================

function Parse-ModList {
    param([string]$ProfilePath)

    $result = @{}
    if ([string]::IsNullOrWhiteSpace($ProfilePath)) { return $result }

    $modlistFile = Join-Path $ProfilePath "modlist.txt"
    if (-not (Test-Path $modlistFile)) { return $result }

    try {
        $lines = Get-Content $modlistFile -Encoding UTF8 -ErrorAction Stop
    } catch { return $result }

    # Filter to mod lines only (start with + or -, not # or *)
    $modLines = @()
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $first = $trimmed.Substring(0, 1)
        if ($first -eq '+' -or $first -eq '-') {
            $modLines += $trimmed
        }
    }

    # Reverse: file's last line = highest priority = priority 1
    [array]::Reverse($modLines)

    $priority = 1
    foreach ($line in $modLines) {
        $enabled  = $line.Substring(0, 1) -eq '+'
        $modName  = $line.Substring(1).Trim()
        # The same mod name should appear only once; first occurrence wins
        if (-not $result.ContainsKey($modName)) {
            $result[$modName] = [PSCustomObject]@{
                Priority = $priority
                Enabled  = $enabled
            }
            $priority++
        }
    }

    return $result
}


# =====================================================================
#  Helper functions (same as CLI version)
# =====================================================================

function Test-IsEspFe {
    param([string]$EspPath)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($EspPath) | Select-Object -First 16
        if ($bytes.Count -lt 10) { return $false }
        return ([byte]($bytes[9]) -band 0x02) -ne 0
    }
    catch { return $false }
}

function Get-NexusUrl {
    param([string]$ModId, [string]$Game)
    if ([string]::IsNullOrWhiteSpace($ModId) -or $ModId -eq "0" -or $ModId -eq "-1") { return "" }
    return "https://www.nexusmods.com/$Game/mods/$ModId"
}

function Get-ModNameFromArchive {
    param([string]$ArchiveName, [string]$ModId)
    if ([string]::IsNullOrWhiteSpace($ArchiveName)) { return "" }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($ArchiveName)
    if (-not [string]::IsNullOrWhiteSpace($ModId) -and $ModId -ne "0") {
        $pattern = "-$ModId-"
        $idx = $name.IndexOf($pattern)
        if ($idx -gt 0) { return $name.Substring(0, $idx) }
    }
    if ($name -match '^(.+?)-\d+') { return $matches[1] }
    return $name
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -lt 1KB)  { return "$Bytes B" }
    if ($Bytes -lt 1MB)  { return "{0:N1} KB" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB)  { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N2} GB" -f ($Bytes / 1GB)
}


# =====================================================================
#  Main analysis
# =====================================================================

function Invoke-Analysis {
    param(
        [string]$ModsPath,
        [string]$ProfilePath,
        [string]$GameSlug,
        [bool]$DoTxt,
        [bool]$DoCsv,
        [bool]$DoHtml
    )

    if (-not (Test-Path $ModsPath)) {
        Write-Log ((Get-Text 'log_error_folder') -f $ModsPath) "#FF6B6B"
        Set-Status (Get-Text 'status_error_folder') 'error'
        return
    }

    Write-Log ((Get-Text 'log_scanning') -f $ModsPath)
    Set-Status (Get-Text 'status_scanning') 'busy'

    # Parse modlist.txt if profile path provided
    $modListData = @{}
    if (-not [string]::IsNullOrWhiteSpace($ProfilePath)) {
        $modListData = Parse-ModList -ProfilePath $ProfilePath
        if ($modListData.Count -gt 0) {
            Write-Log ((Get-Text 'log_modlist_found') -f $modListData.Count) "#A6E22E"
        } else {
            Write-Log (Get-Text 'log_modlist_missing') "#FFCC66"
        }
    } else {
        Write-Log (Get-Text 'log_no_profile') "#888888"
    }

    $allMods = Get-ChildItem $ModsPath -Directory -ErrorAction SilentlyContinue
    $total = $allMods.Count
    if ($total -eq 0) {
        Write-Log ((Get-Text 'log_no_folders') -f $ModsPath) "#FFCC66"
        Set-Status (Get-Text 'status_nothing') 'idle'
        return
    }

    Write-Log ((Get-Text 'log_found_folders') -f $total)
    $results = New-Object System.Collections.ArrayList

    $counter = 0
    foreach ($mod in $allMods) {
        $counter++
        if ($counter % 10 -eq 0 -or $counter -eq $total) {
            $percent = [int](($counter / $total) * 100)
            Set-Progress $percent
            Set-Status ((Get-Text 'status_analyzing') -f $counter, $total) 'busy'
        }

        $metaFile = Join-Path $mod.FullName "meta.ini"
        $instFile = ""; $modId = ""; $version = ""; $gameName = ""
        $hasMeta = Test-Path $metaFile

        if ($hasMeta) {
            $content = Get-Content $metaFile -Raw -ErrorAction SilentlyContinue
            if ($content -match 'installationFile\s*=\s*(.*)')  { $instFile = $matches[1].Trim() }
            if ($content -match '(?m)^\s*modid\s*=\s*(.*)')    { $modId    = $matches[1].Trim() }
            if ($content -match '(?m)^\s*version\s*=\s*(.*)')  { $version  = $matches[1].Trim() }
            if ($content -match '(?m)^\s*gameName\s*=\s*(.*)') { $gameName = $matches[1].Trim() }
        }

        $installType = ""; $reason = ""
        if (-not $hasMeta) {
            $installType = "NoMeta"; $reason = "meta.ini is missing"
        }
        elseif ([string]::IsNullOrWhiteSpace($instFile)) {
            $installType = "NoMeta"; $reason = "installationFile is empty or missing"
        }
        elseif ($instFile -match '[\\/]') {
            $installType = "Manual"; $reason = "installationFile contains a path to an external archive"
        }
        else {
            $installType = "MO2"; $reason = "installationFile is just a filename (installed via MO2)"
        }

        $espFiles = @(); $esmFiles = @(); $eslFiles = @(); $espFeFiles = @()
        $bsaFiles = @(); $ba2Files = @(); $bsaSize = 0L

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
        catch { }

        $hasPlugin = ($espFiles.Count + $esmFiles.Count + $eslFiles.Count) -gt 0

        $nexusName = ""
        if ($installType -eq "Manual") {
            $archiveOnly = Split-Path $instFile -Leaf
            $nexusName = Get-ModNameFromArchive -ArchiveName $archiveOnly -ModId $modId
        }
        elseif ($installType -eq "MO2" -and -not [string]::IsNullOrWhiteSpace($instFile)) {
            $nexusName = Get-ModNameFromArchive -ArchiveName $instFile -ModId $modId
        }

        $nexusUrl = Get-NexusUrl -ModId $modId -Game $GameSlug

        # Look up priority/enabled status from modlist.txt
        $priority = $null; $enabled = $null
        if ($modListData.ContainsKey($mod.Name)) {
            $priority = $modListData[$mod.Name].Priority
            $enabled  = $modListData[$mod.Name].Enabled
        }

        [void]$results.Add([PSCustomObject]@{
            Mod              = $mod.Name
            Priority         = $priority
            Enabled          = $enabled
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
            Plugins          = ($espFiles + $esmFiles + $eslFiles) -join "; "
            BsaCount         = $bsaFiles.Count + $ba2Files.Count
            BsaSize          = $bsaSize
            BsaSizeFormatted = (Format-Size $bsaSize)
            Bsas             = ($bsaFiles + $ba2Files) -join "; "
        })
    }

    $mo2Count    = ($results | Where-Object { $_.InstallType -eq "MO2"    }).Count
    $manualCount = ($results | Where-Object { $_.InstallType -eq "Manual" }).Count
    $noMetaCount = ($results | Where-Object { $_.InstallType -eq "NoMeta" }).Count
    $suspicious  = $results | Where-Object {
        ($_.InstallType -eq "Manual" -or $_.InstallType -eq "NoMeta") -and
        -not $_.HasPlugin -and $_.BsaCount -eq 0
    }

    $hasModlist  = $modListData.Count -gt 0
    $enabledCount  = if ($hasModlist) { ($results | Where-Object { $_.Enabled -eq $true }).Count } else { 0 }
    $disabledCount = if ($hasModlist) { ($results | Where-Object { $_.Enabled -eq $false }).Count } else { 0 }

    Write-Log ""
    Write-Log (Get-Text 'log_stats_header')
    Write-Log ((Get-Text 'log_stats_total')  -f $total)
    Write-Log ((Get-Text 'log_stats_mo2')    -f $mo2Count)
    Write-Log ((Get-Text 'log_stats_manual') -f $manualCount)
    Write-Log ((Get-Text 'log_stats_nometa') -f $noMetaCount)
    Write-Log ((Get-Text 'log_stats_susp')   -f $suspicious.Count)
    if ($hasModlist) {
        Write-Log ((Get-Text 'log_stats_enabled')  -f $enabledCount)
        Write-Log ((Get-Text 'log_stats_disabled') -f $disabledCount)
    }
    Write-Log ""

    if ($suspicious.Count -gt 0) {
        Write-Log (Get-Text 'log_susp_header') "#FF6B6B"
        foreach ($s in ($suspicious | Sort-Object Mod)) {
            Write-Log "  - $($s.Mod)" "#FF6B6B"
        }
        Write-Log ""
    }

    if ($DoTxt) {
        Write-TxtReport -Results $results -ModsPath $ModsPath `
            -Mo2Count $mo2Count -ManualCount $manualCount -NoMetaCount $noMetaCount `
            -Suspicious $suspicious -Total $total -HasModlist $hasModlist `
            -EnabledCount $enabledCount -DisabledCount $disabledCount `
            -Path (Join-Path $outputDir "ManualMods-Report.txt")
        Write-Log (Get-Text 'log_txt_saved') "#A6E22E"
    }

    if ($DoCsv) {
        $csvData = $results | Select-Object Priority, Enabled, Mod, InstallType, ModId, Version, NexusName, NexusUrl,
            HasPlugin, EspCount, EsmCount, EslCount, EspFeCount, BsaCount, BsaSizeFormatted,
            Plugins, Bsas, InstallationFile, Reason | Sort-Object @{Expression={if ($null -eq $_.Priority) { [int]::MaxValue } else { $_.Priority }}}, Mod
        $csvData | Export-Csv -Path (Join-Path $outputDir "ManualMods-Report.csv") -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log (Get-Text 'log_csv_saved') "#A6E22E"
    }

    if ($DoHtml) {
        Write-HtmlReport -Results $results -ModsPath $ModsPath `
            -Mo2Count $mo2Count -ManualCount $manualCount -NoMetaCount $noMetaCount `
            -Suspicious $suspicious -Total $total -HasModlist $hasModlist `
            -EnabledCount $enabledCount -DisabledCount $disabledCount `
            -Path (Join-Path $outputDir "ManualMods-Report.html")
        Write-Log (Get-Text 'log_html_saved') "#A6E22E"
    }

    Write-Log ""
    Write-Log ((Get-Text 'log_reports_to') -f $outputDir)
    Set-Status ((Get-Text 'status_done') -f $suspicious.Count) 'done'
    Set-Progress 100

    $openFolderButton.Dispatcher.Invoke([Action]{ $openFolderButton.IsEnabled = $true })
}


# =====================================================================
#  Report writers
# =====================================================================

function Write-TxtReport {
    param($Results, $ModsPath, $Mo2Count, $ManualCount, $NoMetaCount, $Suspicious, $Total, $HasModlist, $EnabledCount, $DisabledCount, $Path)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  MOD ORGANIZER 2 - INSTALLED MODS REPORT")
    [void]$sb.AppendLine("  Date:   $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("  Folder: $ModsPath")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("STATISTICS:")
    [void]$sb.AppendLine("  Total mod folders:                       $Total")
    [void]$sb.AppendLine("  Installed via MO2 (built-in downloader): $Mo2Count")
    [void]$sb.AppendLine("  Installed manually (external archive):   $ManualCount")
    [void]$sb.AppendLine("  No meta.ini or no installationFile:      $NoMetaCount")
    [void]$sb.AppendLine("  SUSPICIOUS (no plugins and no BSA):      $($Suspicious.Count)")
    if ($HasModlist) {
        [void]$sb.AppendLine("  Enabled in profile:                      $EnabledCount")
        [void]$sb.AppendLine("  Disabled in profile:                     $DisabledCount")
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  SUSPICIOUS MODS")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    if ($Suspicious.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else {
        foreach ($item in ($Suspicious | Sort-Object Mod)) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Install type: $($item.InstallType)")
            [void]$sb.AppendLine("    Reason:       $($item.Reason)")
            if ($null -ne $item.Priority) {
                $en = if ($item.Enabled) { "yes" } else { "no" }
                [void]$sb.AppendLine("    Priority:     $($item.Priority) (enabled: $en)")
            }
            if ($item.NexusUrl) { [void]$sb.AppendLine("    Nexus:        $($item.NexusUrl)") }
            if ($item.InstallationFile) { [void]$sb.AppendLine("    Archive:      $($item.InstallationFile)") }
            [void]$sb.AppendLine("")
        }
    }

    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  LIST 1: MODS INSTALLED FROM EXTERNAL ARCHIVES")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    $manuals = $Results | Where-Object { $_.InstallType -eq "Manual" } | Sort-Object Mod
    if ($manuals.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else {
        foreach ($item in $manuals) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Archive:  $($item.InstallationFile)")
            if ($null -ne $item.Priority) {
                $en = if ($item.Enabled) { "yes" } else { "no" }
                [void]$sb.AppendLine("    Priority: $($item.Priority) (enabled: $en)")
            }
            if ($item.ModId)    { [void]$sb.AppendLine("    Nexus ID: $($item.ModId)") }
            if ($item.Version)  { [void]$sb.AppendLine("    Version:  $($item.Version)") }
            if ($item.NexusUrl) { [void]$sb.AppendLine("    URL:      $($item.NexusUrl)") }
            $line = "    Plugins:  "
            if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount   -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount   -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount   -gt 0) { $parts += "$($item.EslCount) ESL" }
                if ($item.EspFeCount -gt 0) { $parts += "($($item.EspFeCount) of which are ESPFE)" }
                $line += ($parts -join ", ")
            } else { $line += "NONE" }
            [void]$sb.AppendLine($line)
            if ($item.BsaCount -gt 0) { [void]$sb.AppendLine("    BSA/BA2:  $($item.BsaCount) file(s) ($($item.BsaSizeFormatted))") }
            [void]$sb.AppendLine("")
        }
    }

    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("  LIST 2: MODS WITHOUT META.INI OR INSTALLATIONFILE")
    [void]$sb.AppendLine("================================================================")
    [void]$sb.AppendLine("")
    $noMetas = $Results | Where-Object { $_.InstallType -eq "NoMeta" } | Sort-Object Mod
    if ($noMetas.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else {
        foreach ($item in $noMetas) {
            [void]$sb.AppendLine("  [$($item.Mod)]")
            [void]$sb.AppendLine("    Reason:   $($item.Reason)")
            if ($null -ne $item.Priority) {
                $en = if ($item.Enabled) { "yes" } else { "no" }
                [void]$sb.AppendLine("    Priority: $($item.Priority) (enabled: $en)")
            }
            $line = "    Plugins:  "
            if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount -gt 0) { $parts += "$($item.EslCount) ESL" }
                $line += ($parts -join ", ")
            } else { $line += "NONE" }
            [void]$sb.AppendLine($line)
            if ($item.BsaCount -gt 0) { [void]$sb.AppendLine("    BSA/BA2:  $($item.BsaCount) file(s) ($($item.BsaSizeFormatted))") }
            [void]$sb.AppendLine("")
        }
    }

    if ($HasModlist) {
        [void]$sb.AppendLine("================================================================")
        [void]$sb.AppendLine("  LOAD ORDER (from modlist.txt, priority 1 = top of MO2 UI)")
        [void]$sb.AppendLine("================================================================")
        [void]$sb.AppendLine("")
        $sortedByPriority = $Results | Where-Object { $null -ne $_.Priority } | Sort-Object Priority
        foreach ($item in $sortedByPriority) {
            $en = if ($item.Enabled) { "[X]" } else { "[ ]" }
            $prio = "{0,4}" -f $item.Priority
            [void]$sb.AppendLine("  $prio. $en $($item.Mod)")
        }
        [void]$sb.AppendLine("")
    }

    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), $utf8Bom)
}

function Write-HtmlReport {
    param($Results, $ModsPath, $Mo2Count, $ManualCount, $NoMetaCount, $Suspicious, $Total, $HasModlist, $EnabledCount, $DisabledCount, $Path)

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine("<!DOCTYPE html>")
    [void]$html.AppendLine("<html lang='en'><head><meta charset='UTF-8'><title>MO2 Mods Report</title>")
    [void]$html.AppendLine("<style>")
    [void]$html.AppendLine("body{font-family:'Segoe UI',Tahoma,sans-serif;background:#1e1e1e;color:#d4d4d4;padding:20px}")
    [void]$html.AppendLine("h1{color:#569cd6}h2{color:#4ec9b0;border-bottom:2px solid #4ec9b0;padding-bottom:5px;margin-top:30px}")
    [void]$html.AppendLine(".stats{background:#252526;padding:15px;border-radius:6px;margin-bottom:20px}.stats div{margin:4px 0}")
    [void]$html.AppendLine(".badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:.85em;font-weight:bold}")
    [void]$html.AppendLine(".bg-mo2{background:#2d7a2d;color:white}.bg-manual{background:#b58900;color:white}")
    [void]$html.AppendLine(".bg-nometa{background:#cb4b16;color:white}.bg-danger{background:#dc322f;color:white}")
    [void]$html.AppendLine(".bg-on{background:#2d7a2d;color:white}.bg-off{background:#555;color:#ccc}")
    [void]$html.AppendLine("table{border-collapse:collapse;width:100%;background:#252526;margin-bottom:30px}")
    [void]$html.AppendLine("th,td{padding:8px 12px;text-align:left;border-bottom:1px solid #3e3e42;vertical-align:top}")
    [void]$html.AppendLine("th{background:#2d2d30;color:#569cd6;cursor:pointer;user-select:none}th:hover{background:#3e3e42}")
    [void]$html.AppendLine("tr:hover{background:#2a2a2d}a{color:#569cd6;text-decoration:none}a:hover{text-decoration:underline}")
    [void]$html.AppendLine(".filter-box{margin-bottom:10px}.filter-box input{background:#3c3c3c;color:#d4d4d4;border:1px solid #555;padding:6px 10px;border-radius:4px;width:300px}")
    [void]$html.AppendLine(".small{font-size:.85em;color:#888}.danger-row{background:#3a1a1a !important}.disabled-row{opacity:0.5}")
    [void]$html.AppendLine(".prio{color:#888;font-family:Consolas,monospace}")
    [void]$html.AppendLine("</style></head><body>")

    [void]$html.AppendLine("<h1>Mod Organizer 2 - mods report</h1>")
    [void]$html.AppendLine("<div class='small'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>Folder: $ModsPath</div>")

    [void]$html.AppendLine("<div class='stats'>")
    [void]$html.AppendLine("<h2 style='margin-top:0;border:none;padding:0'>Statistics</h2>")
    [void]$html.AppendLine("<div>Total mods: <b>$Total</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-mo2'>MO2</span> installed via MO2: <b>$Mo2Count</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-manual'>Manual</span> installed manually: <b>$ManualCount</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-nometa'>NoMeta</span> no meta.ini: <b>$NoMetaCount</b></div>")
    [void]$html.AppendLine("<div><span class='badge bg-danger'>!</span> Suspicious: <b>$($Suspicious.Count)</b></div>")
    if ($HasModlist) {
        [void]$html.AppendLine("<div><span class='badge bg-on'>ON</span> enabled in profile: <b>$EnabledCount</b></div>")
        [void]$html.AppendLine("<div><span class='badge bg-off'>OFF</span> disabled in profile: <b>$DisabledCount</b></div>")
    }
    [void]$html.AppendLine("</div>")

    function Write-Tbl {
        param($Sb, $Title, $Items, $TableId, $Highlight, $HasModlist, [string]$DefaultSort = 'Mod')
        [void]$Sb.AppendLine("<h2>$Title (total: $($Items.Count))</h2>")
        if ($Items.Count -eq 0) { [void]$Sb.AppendLine("<p class='small'>(none)</p>"); return }
        [void]$Sb.AppendLine("<div class='filter-box'><input type='text' placeholder='Filter by name...' onkeyup=""filterTable('$TableId',this.value)""></div>")
        [void]$Sb.AppendLine("<table id='$TableId'><thead><tr>")
        $colIdx = 0
        if ($HasModlist) {
            [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx,'num')"">#</th>"); $colIdx++
            [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">On/Off</th>"); $colIdx++
        }
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">Mod</th>"); $colIdx++
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">Type</th>"); $colIdx++
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">Nexus</th>"); $colIdx++
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">Version</th>"); $colIdx++
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">Plugins</th>"); $colIdx++
        [void]$Sb.AppendLine("<th onclick=""sortTable('$TableId',$colIdx)"">BSA/BA2</th>"); $colIdx++
        [void]$Sb.AppendLine("<th>Archive / Reason</th>")
        [void]$Sb.AppendLine("</tr></thead><tbody>")

        # Choose ordering: by priority if available and requested, else by name
        $ordered = if ($DefaultSort -eq 'Priority' -and $HasModlist) {
            $Items | Sort-Object @{Expression={if ($null -eq $_.Priority) { [int]::MaxValue } else { $_.Priority }}}, Mod
        } else {
            $Items | Sort-Object Mod
        }

        foreach ($item in $ordered) {
            $rowClasses = @()
            if ($Highlight -and -not $item.HasPlugin -and $item.BsaCount -eq 0) { $rowClasses += 'danger-row' }
            if ($HasModlist -and $item.Enabled -eq $false) { $rowClasses += 'disabled-row' }
            $rowAttr = if ($rowClasses.Count -gt 0) { " class='$($rowClasses -join " ")'" } else { "" }

            $typeBadge = switch ($item.InstallType) {
                "MO2"    { "<span class='badge bg-mo2'>MO2</span>" }
                "Manual" { "<span class='badge bg-manual'>Manual</span>" }
                "NoMeta" { "<span class='badge bg-nometa'>NoMeta</span>" }
            }
            $nexusCell = if ($item.NexusUrl) { "<a href='$($item.NexusUrl)' target='_blank'>$($item.ModId)</a>" }
                         elseif ($item.ModId) { $item.ModId } else { "-" }
            $pluginCell = if ($item.HasPlugin) {
                $parts = @()
                if ($item.EspCount   -gt 0) { $parts += "$($item.EspCount) ESP" }
                if ($item.EsmCount   -gt 0) { $parts += "$($item.EsmCount) ESM" }
                if ($item.EslCount   -gt 0) { $parts += "$($item.EslCount) ESL" }
                if ($item.EspFeCount -gt 0) { $parts += "<span class='small'>($($item.EspFeCount) ESPFE)</span>" }
                $parts -join ", "
            } else { "<span class='badge bg-danger'>NONE</span>" }
            $bsaCell = if ($item.BsaCount -gt 0) { "$($item.BsaCount) ($($item.BsaSizeFormatted))" } else { "-" }
            $reasonCell = if ($item.InstallType -eq "Manual") {
                "<span class='small'>$([System.Web.HttpUtility]::HtmlEncode($item.InstallationFile))</span>"
            } else {
                "<span class='small'>$([System.Web.HttpUtility]::HtmlEncode($item.Reason))</span>"
            }
            $modName = [System.Web.HttpUtility]::HtmlEncode($item.Mod)
            $version = [System.Web.HttpUtility]::HtmlEncode($item.Version)

            $row = "<tr$rowAttr>"
            if ($HasModlist) {
                $prioCell = if ($null -ne $item.Priority) { "<span class='prio'>$($item.Priority)</span>" } else { "<span class='small'>-</span>" }
                $onOffCell = if ($item.Enabled -eq $true)  { "<span class='badge bg-on'>ON</span>" }
                              elseif ($item.Enabled -eq $false) { "<span class='badge bg-off'>OFF</span>" }
                              else { "<span class='small'>-</span>" }
                $row += "<td>$prioCell</td><td>$onOffCell</td>"
            }
            $row += "<td>$modName</td><td>$typeBadge</td><td>$nexusCell</td><td>$version</td><td>$pluginCell</td><td>$bsaCell</td><td>$reasonCell</td></tr>"
            [void]$Sb.AppendLine($row)
        }
        [void]$Sb.AppendLine("</tbody></table>")
    }

    Write-Tbl $html "Suspicious mods" $Suspicious "tbl-susp" $true $HasModlist 'Mod'
    Write-Tbl $html "Installed manually (external archive)" ($Results | Where-Object { $_.InstallType -eq "Manual" }) "tbl-manual" $true $HasModlist 'Mod'
    Write-Tbl $html "No meta.ini or no installationFile" ($Results | Where-Object { $_.InstallType -eq "NoMeta" }) "tbl-nometa" $true $HasModlist 'Mod'

    # Single "All mods" table - sorted by load order if modlist.txt is available,
    # otherwise alphabetical. User can re-sort by clicking column headers.
    $allModsTitle = if ($HasModlist) {
        "All mods (sorted by load order, priority 1 = top of MO2 UI; click any header to re-sort)"
    } else {
        "All mods (alphabetical; click any header to re-sort)"
    }
    $allModsSort = if ($HasModlist) { 'Priority' } else { 'Mod' }
    Write-Tbl $html $allModsTitle $Results "tbl-all" $false $HasModlist $allModsSort

    [void]$html.AppendLine(@"
<script>
function sortTable(tableId,col,type){var t=document.getElementById(tableId),b=t.tBodies[0],r=Array.from(b.rows);
var asc=t.getAttribute('data-sort-col')!=col||t.getAttribute('data-sort-dir')=='desc';
r.sort(function(a,bb){var x=a.cells[col].innerText.trim().toLowerCase(),y=bb.cells[col].innerText.trim().toLowerCase();
if(type=='num'){var nx=parseInt(x),ny=parseInt(y);if(isNaN(nx))nx=Number.MAX_SAFE_INTEGER;if(isNaN(ny))ny=Number.MAX_SAFE_INTEGER;return asc?nx-ny:ny-nx;}
var nx=parseFloat(x),ny=parseFloat(y);if(!isNaN(nx)&&!isNaN(ny))return asc?nx-ny:ny-nx;
return asc?x.localeCompare(y):y.localeCompare(x);});r.forEach(function(x){b.appendChild(x);});
t.setAttribute('data-sort-col',col);t.setAttribute('data-sort-dir',asc?'asc':'desc');}
function filterTable(tableId,q){q=q.toLowerCase();var rows=document.getElementById(tableId).tBodies[0].rows;
for(var i=0;i<rows.length;i++){var nameCell=rows[i].cells[rows[i].cells.length>=9?2:0];rows[i].style.display=nameCell.innerText.toLowerCase().includes(q)?'':'none';}}
</script>
"@)

    [void]$html.AppendLine("</body></html>")

    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($Path, $html.ToString(), $utf8Bom)
}


# =====================================================================
#  Run button
# =====================================================================

$runButton.Add_Click({
    $modsPath    = $modsPathBox.Text.Trim()
    $profilePath = $profilePathBox.Text.Trim()
    $gameSlug    = $gameSlugBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($modsPath)) {
        [System.Windows.MessageBox]::Show((Get-Text 'msg_missing_path'), (Get-Text 'msg_missing_path_t'), "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($gameSlug)) { $gameSlug = "skyrimspecialedition" }

    Persist-Settings

    $logBox.Clear()
    $progress.Value = 0
    $runButton.IsEnabled = $false
    $openFolderButton.IsEnabled = $false

    $doTxt  = $txtCheck.IsChecked
    $doCsv  = $csvCheck.IsChecked
    $doHtml = $htmlCheck.IsChecked

    try {
        Invoke-Analysis -ModsPath $modsPath -ProfilePath $profilePath -GameSlug $gameSlug -DoTxt $doTxt -DoCsv $doCsv -DoHtml $doHtml
    }
    catch {
        Write-Log ((Get-Text 'log_error_generic') -f $_) "#FF6B6B"
        Set-Status (Get-Text 'status_failed') 'error'
    }
    finally {
        $runButton.IsEnabled = $true
    }
})


# =====================================================================
#  Save settings on close, then show window
# =====================================================================

$window.Add_Closing({ Persist-Settings })

[void]$window.ShowDialog()
