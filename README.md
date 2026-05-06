# MO2-Manual-Install-Auditor
A PowerShell diagnostic script that scans your Mod Organizer 2 mods folder, identifies mods installed manually (vs through MO2's built-in downloader), and generates TXT/CSV/HTML reports with plugin and BSA info.
================================================================
  MO2 Manual Install Auditor - v2.0
================================================================

A diagnostic tool for Mod Organizer 2 that finds mods installed
manually (vs through MO2's built-in downloader), reports their
load order, plugin counts, and BSA archives, and flags broken
manual installs.


----------------------------------------------------------------
  CONTENTS OF THIS ARCHIVE
----------------------------------------------------------------

  MO2ManualInstallAuditor.exe   - Main GUI application
  Find-ManualMods-GUI.ps1       - GUI source code (PowerShell)
  Find-ManualMods.ps1           - Original CLI version
  README.txt                    - This file


----------------------------------------------------------------
  WHAT IS NEW IN v2.0
----------------------------------------------------------------

  - Graphical interface (no more editing scripts)
  - Compiled .exe - just double-click to run
  - Settings are remembered between sessions
  - English / Russian interface (more languages can be added)
  - Load order analysis from modlist.txt (priority + on/off)
  - HTML report shows load order column with on/off badges,
    disabled mods are visually dimmed
  - All settings stored in:
      %APPDATA%\MO2ManualInstallAuditor\settings.json


----------------------------------------------------------------
  HOW TO USE
----------------------------------------------------------------

  1. Run MO2ManualInstallAuditor.exe (double-click)

  2. Click "Browse..." next to "Path to MO2 mods folder" and
     select the mods folder of your MO2 instance.
     Example: D:\MO2\Skyrim Special Edition\mods

  3. (Optional) Click "Browse..." next to the profile folder
     field and select an MO2 profile folder. This enables
     load order analysis.
     Example: D:\MO2\Skyrim Special Edition\profiles\Default

  4. (Optional) Change Nexus game slug if you are not using
     Skyrim SE. Examples: fallout4, starfield, oblivion.

  5. Choose which reports to generate (TXT / CSV / HTML).

  6. Click "Run Analysis".

  7. When finished, click "Open Reports Folder".
     The reports are saved next to the .exe file.

  Open ManualMods-Report.html in your browser - it is the most
  convenient way to explore the results. Every column header is
  clickable to re-sort. Suspicious mods are highlighted in red.


----------------------------------------------------------------
  HOW TO VERIFY THE .EXE IS SAFE
----------------------------------------------------------------

  This is an unsigned executable compiled from a PowerShell
  script. Windows SmartScreen will likely show a warning on
  first run ("Windows protected your PC"). This is normal for
  any unsigned executable on Nexus - the warning means Windows
  has not seen this file from many users yet, not that it is
  malicious.

  You have several ways to verify the file is safe:

  ----- METHOD 1: Read the source code -----

  The full source is in Find-ManualMods-GUI.ps1, included in
  this archive. Open it in any text editor (Notepad, Notepad++,
  VS Code) and read it. The script:

    - Only READS files in the folder you specify
    - Writes report files (TXT, CSV, HTML) next to itself
    - Writes settings JSON in %APPDATA%
    - Makes NO network requests
    - Does NOT modify, move, or delete any of your mods

  No obfuscation, no encoded strings, no encrypted blobs - the
  whole logic is plain text PowerShell.

  ----- METHOD 2: Run the script directly instead of the .exe -----

  If you do not trust the .exe, you can run the PowerShell
  script directly - it does the exact same thing:

    1. Open PowerShell in the folder containing the script
    2. Unblock the file:
         Unblock-File .\Find-ManualMods-GUI.ps1
    3. If execution policy blocks scripts, run once:
         Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
    4. Run:
         .\Find-ManualMods-GUI.ps1

  ----- METHOD 3: Scan with VirusTotal -----

  Upload MO2ManualInstallAuditor.exe to:
    https://www.virustotal.com

  The site scans the file with 60+ antivirus engines and shows
  results from all of them. A few false positives (1-3 out of
  60) are common and expected for ps2exe-compiled files - this
  is a known limitation of the tool, not a sign of malware.
  Many legitimate Nexus mods have similar VT results.

  ----- METHOD 4: Compile the .exe yourself -----

  If you want a guaranteed-safe build, compile it from the
  source script yourself:

    1. Open PowerShell as your normal user
    2. Install ps2exe (one-time):
         Install-Module -Name ps2exe -Scope CurrentUser -Force
    3. Compile:
         Invoke-PS2EXE .\Find-ManualMods-GUI.ps1 .\MyBuild.exe -noConsole

  The result will be functionally identical to the .exe shipped
  in this archive, but built from source you reviewed.

  ----- METHOD 5: Run from a sandbox -----

  Tools like Sandboxie, Windows Sandbox (built into Win 10/11
  Pro), or a virtual machine let you run the .exe in an
  isolated environment to observe what it does.


----------------------------------------------------------------
  WHAT IF SMARTSCREEN BLOCKS THE .EXE
----------------------------------------------------------------

  When you first run the .exe, Windows may show:

    "Windows protected your PC"
    Microsoft Defender SmartScreen prevented an unrecognized
    app from starting...

  To run anyway:
    1. Click "More info" (small text on the dialog)
    2. Click the "Run anyway" button that appears

  Or to remove the "downloaded from internet" mark entirely:
    1. Right-click the .exe -> Properties
    2. Bottom of "General" tab - check "Unblock"
    3. Click OK


----------------------------------------------------------------
  WHAT IF MY ANTIVIRUS FLAGS THE .EXE
----------------------------------------------------------------

  Some antivirus engines flag ps2exe-compiled executables as
  suspicious by heuristic, because malware sometimes uses the
  same packing tool. This is a false positive.

  Options:
    - Add an exception for the file in your antivirus
    - Use the .ps1 script directly (no compilation involved)
    - Compile the .exe yourself (Method 4 above)
    - Verify on VirusTotal first

  If you encounter a serious flag from a reputable AV (not just
  generic "PUA/ps2exe heuristic"), please report it in the
  comments and I will investigate.


----------------------------------------------------------------
  WHERE TO FIND SETTINGS AND OUTPUT
----------------------------------------------------------------

  Settings file (paths, language, checkboxes):
    %APPDATA%\MO2ManualInstallAuditor\settings.json

  Report output (next to the .exe by default):
    ManualMods-Report.txt
    ManualMods-Report.csv
    ManualMods-Report.html

  To reset settings: delete the settings.json file.
  To uninstall: just delete the .exe and the settings folder.


----------------------------------------------------------------
  REQUIREMENTS
----------------------------------------------------------------

  - Windows 10 or 11
  - PowerShell 5.1+ (built into Windows, nothing to install)
  - Mod Organizer 2 (any recent version)


----------------------------------------------------------------
  LICENSE & PERMISSIONS
----------------------------------------------------------------

  See the mod page on Nexus for full permission details.

  Short version: free to use, modify, fork, port to other
  games (with credit), share. No monetization.


================================================================
================================================================
  RUSSIAN VERSION / РУССКАЯ ВЕРСИЯ
================================================================
================================================================


----------------------------------------------------------------
  СОДЕРЖИМОЕ АРХИВА
----------------------------------------------------------------

  MO2ManualInstallAuditor.exe   - Основное GUI-приложение
  Find-ManualMods-GUI.ps1       - Исходный код GUI (PowerShell)
  Find-ManualMods.ps1           - Оригинальная CLI-версия
  README.txt                    - Этот файл


----------------------------------------------------------------
  ЧТО НОВОГО В v2.0
----------------------------------------------------------------

  - Графический интерфейс (больше не нужно править скрипты)
  - Скомпилированный .exe - двойной клик и запуск
  - Настройки сохраняются между запусками
  - Английский / русский интерфейс (можно добавить другие языки)
  - Анализ порядка загрузки из modlist.txt (приоритет + вкл/выкл)
  - HTML-отчёт показывает столбец порядка загрузки и значки
    включения/отключения, выключенные моды визуально затемнены
  - Все настройки хранятся в:
      %APPDATA%\MO2ManualInstallAuditor\settings.json


----------------------------------------------------------------
  КАК ПОЛЬЗОВАТЬСЯ
----------------------------------------------------------------

  1. Запустите MO2ManualInstallAuditor.exe (двойной клик).

  2. Нажмите "Обзор..." рядом с полем "Путь к папке mods" и
     выберите папку mods вашего инстанса MO2.
     Например: D:\MO2\Skyrim Special Edition\mods

  3. (Опционально) Нажмите "Обзор..." рядом с полем папки
     профиля и выберите папку профиля MO2. Это включит анализ
     порядка загрузки модов.
     Например: D:\MO2\Skyrim Special Edition\profiles\Default

  4. (Опционально) Измените Nexus game slug, если используете
     не Skyrim SE. Примеры: fallout4, starfield, oblivion.

  5. Выберите форматы отчётов (TXT / CSV / HTML).

  6. Нажмите "Запустить анализ".

  7. По завершении нажмите "Открыть папку с отчётами".
     Отчёты сохраняются рядом с .exe файлом.

  Откройте ManualMods-Report.html в браузере - это самый
  удобный способ изучить результаты. Клик по заголовку любого
  столбца меняет сортировку. Подозрительные моды подсвечены
  красным.


----------------------------------------------------------------
  КАК ПРОВЕРИТЬ БЕЗОПАСНОСТЬ .EXE
----------------------------------------------------------------

  Это неподписанный exe, скомпилированный из PowerShell-скрипта.
  Windows SmartScreen скорее всего покажет предупреждение при
  первом запуске ("Система Windows защитила ваш компьютер").
  Это нормально для любых неподписанных exe на Nexus - просто
  Windows ещё не видела этот файл у многих пользователей, это
  не значит, что файл вредоносный.

  Есть несколько способов проверить, что файл безопасен:

  ----- СПОСОБ 1: Прочитать исходный код -----

  Полный исходник лежит в архиве - Find-ManualMods-GUI.ps1.
  Откройте его в любом текстовом редакторе (Блокнот, Notepad++,
  VS Code) и прочитайте. Скрипт:

    - Только ЧИТАЕТ файлы в указанной вами папке
    - Создаёт файлы отчётов (TXT, CSV, HTML) рядом с собой
    - Создаёт JSON с настройками в %APPDATA%
    - НЕ обращается к интернету
    - НЕ изменяет, не перемещает и не удаляет ваши моды

  Никакой обфускации, закодированных строк или зашифрованных
  блоков - вся логика - это обычный текст PowerShell.

  ----- СПОСОБ 2: Запустить скрипт напрямую вместо .exe -----

  Если не доверяете .exe, можно запустить PowerShell-скрипт
  напрямую - результат будет тот же:

    1. Откройте PowerShell в папке со скриптом
    2. Разблокируйте файл:
         Unblock-File .\Find-ManualMods-GUI.ps1
    3. Если политика выполнения блокирует скрипты, выполните
       один раз:
         Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
    4. Запустите:
         .\Find-ManualMods-GUI.ps1

  ----- СПОСОБ 3: Проверить через VirusTotal -----

  Загрузите MO2ManualInstallAuditor.exe на:
    https://www.virustotal.com

  Сайт прогонит файл через 60+ антивирусов и покажет результаты
  каждого. Несколько ложных срабатываний (1-3 из 60) - это
  норма для ps2exe-компилированных файлов, это известное
  ограничение инструмента, а не признак вируса. У многих
  легитимных модов на Nexus похожие результаты на VT.

  ----- СПОСОБ 4: Скомпилировать .exe самостоятельно -----

  Если хотите гарантированно безопасную сборку - скомпилируйте
  из исходника сами:

    1. Откройте PowerShell от своего пользователя
    2. Установите ps2exe (один раз):
         Install-Module -Name ps2exe -Scope CurrentUser -Force
    3. Скомпилируйте:
         Invoke-PS2EXE .\Find-ManualMods-GUI.ps1 .\MyBuild.exe -noConsole

  Результат будет функционально идентичен .exe из архива, но
  собран из исходника, который вы прочитали.

  ----- СПОСОБ 5: Запустить в песочнице -----

  Sandboxie, Windows Sandbox (встроена в Win 10/11 Pro) или
  виртуальная машина позволяют запустить .exe в изолированной
  среде, чтобы посмотреть, что он делает.


----------------------------------------------------------------
  ЕСЛИ SMARTSCREEN БЛОКИРУЕТ .EXE
----------------------------------------------------------------

  При первом запуске .exe Windows может показать:

    "Система Windows защитила ваш компьютер"
    Фильтр SmartScreen в Microsoft Defender предотвратил запуск
    неопознанного приложения...

  Чтобы запустить:
    1. Нажмите "Подробнее" (мелкий текст в окне)
    2. Нажмите появившуюся кнопку "Выполнить в любом случае"

  Или чтобы вообще снять метку "скачано из интернета":
    1. Правый клик по .exe -> Свойства
    2. Внизу вкладки "Общие" - поставьте галку "Разблокировать"
    3. Нажмите OK


----------------------------------------------------------------
  ЕСЛИ АНТИВИРУС ПОМЕЧАЕТ .EXE
----------------------------------------------------------------

  Некоторые антивирусы помечают ps2exe-компилированные файлы
  как подозрительные по эвристике, потому что иногда вирусы
  используют тот же инструмент упаковки. Это ложное срабатывание.

  Варианты:
    - Добавить файл в исключения антивируса
    - Использовать .ps1-скрипт напрямую (без компиляции)
    - Скомпилировать .exe самостоятельно (способ 4 выше)
    - Проверить через VirusTotal перед запуском

  Если столкнётесь с серьёзным срабатыванием от уважаемого
  антивируса (не просто общим "PUA/ps2exe heuristic") - сообщите
  в комментариях, я разберусь.


----------------------------------------------------------------
  ГДЕ ХРАНЯТСЯ НАСТРОЙКИ И ОТЧЁТЫ
----------------------------------------------------------------

  Файл настроек (пути, язык, чекбоксы):
    %APPDATA%\MO2ManualInstallAuditor\settings.json

  Файлы отчётов (по умолчанию рядом с .exe):
    ManualMods-Report.txt
    ManualMods-Report.csv
    ManualMods-Report.html

  Чтобы сбросить настройки - удалите файл settings.json.
  Чтобы удалить программу - просто удалите .exe и папку настроек.


----------------------------------------------------------------
  ТРЕБОВАНИЯ
----------------------------------------------------------------

  - Windows 10 или 11
  - PowerShell 5.1+ (встроен в Windows, ничего ставить не надо)
  - Mod Organizer 2 (любая актуальная версия)


----------------------------------------------------------------
  ЛИЦЕНЗИЯ И РАЗРЕШЕНИЯ
----------------------------------------------------------------

  Полные условия - на странице мода на Nexus.

  Кратко: можно использовать, модифицировать, форкать,
  портировать на другие игры (с указанием авторства),
  делиться. Без монетизации.
