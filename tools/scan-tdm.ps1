# ============================================================================
#  scan-tdm.ps1  -  BEZOPASNYY "osmotrshchik" prilozheniya TDM Messenger
#  (kommentarii po-russki nizhe; fayl sokhranen v UTF-8 s BOM)
# ============================================================================
#
#  ЧТО ДЕЛАЕТ:
#    Находит папку, куда установлен TDM Messenger, и заглядывает внутрь:
#    из чего программа собрана и есть ли у неё зацепки для подключения
#    других программ (боты, api, токены, веб-адреса).
#    Результат складывает в текстовый файл на Рабочем столе.
#
#  ЧЕГО НЕ ДЕЛАЕТ (важно):
#    - НЕ запускает TDM и никакие другие программы
#    - НЕ меняет и не удаляет ни одного файла
#    - НЕ отправляет НИЧЕГО в интернет
#    - НЕ собирает ваши пароли: длинные "секретные" строки прячет звёздочками
#    - НЕ требует прав администратора
#
#  КАК ЗАПУСТИТЬ:
#    1) Нажми "Пуск", напечатай  PowerShell , открой "Windows PowerShell".
#    2) Напечатай команду (подставь свой путь к файлу) и нажми Enter:
#       powershell -ExecutionPolicy Bypass -File "C:\Users\akritskiy\Desktop\scan-tdm.ps1"
#
#    Если сканер не нашёл TDM сам - укажи папку вручную:
#       powershell -ExecutionPolicy Bypass -File "C:\Users\akritskiy\Desktop\scan-tdm.ps1" -Path "C:\путь\к\TDM"
#
#  КОГДА ЗАКОНЧИТ: на Рабочем столе появится TDM_scan_report.txt
#  Открой его, скопируй текст и пришли мне в чат.
# ============================================================================

param(
    [string]$Path = "",
    # Сюда попадут "хвосты" пути, если в названии папки есть пробелы
    # (например "C:\Program Files\TDM Messenger") и Windows разорвал его на части.
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Rest
)

# Если путь разорвало пробелами - склеиваем обратно.
if ($Rest -and @($Rest).Count -gt 0 -and -not [string]::IsNullOrEmpty($Path)) {
    if (-not (Test-Path -LiteralPath $Path)) {
        $glued = ($Path + ' ' + (@($Rest) -join ' ')).Trim()
        if (Test-Path -LiteralPath $glued) { $Path = $glued }
    }
}

# --- Куда сохраняем отчёт (учитывает Рабочий стол, перенесённый в OneDrive) ---
$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrEmpty($desktop)) { $desktop = Join-Path $env:USERPROFILE 'Desktop' }
$reportPath = Join-Path $desktop 'TDM_scan_report.txt'

$report = New-Object System.Collections.Generic.List[string]

function Say([string]$line) {
    Write-Host $line
    $null = $report.Add($line)
}

function SaveReport() {
    try {
        $report | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host ""
        Write-Host "Отчёт сохранён: $reportPath"
    } catch {
        Write-Host "Не удалось сохранить отчёт: $($_.Exception.Message)"
    }
}

# Прячем длинные "секретные" строки (32+ символа подряд), чтобы возможные
# пароли и ключи не попали в отчёт.
function Mask([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return "" }
    return [System.Text.RegularExpressions.Regex]::Replace($text, '[A-Za-z0-9_\+/=]{32,}', '***СКРЫТО***')
}

Say "============================================================"
Say " ОСМОТР TDM MESSENGER - отчёт"
Say " Дата: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Say " Система: $([Environment]::OSVersion.VersionString)"
Say " PowerShell: $($PSVersionTable.PSVersion)"
Say "============================================================"
Say ""

$candidates = New-Object System.Collections.Generic.List[string]

# ---------------------------------------------------------------------------
# ШАГ 1. Ищем папку приложения TDM тремя способами
# ---------------------------------------------------------------------------
Say "[1] Ищу, где установлен TDM"

# --- Способ 1: путь, указанный вручную ---
if (-not [string]::IsNullOrEmpty($Path)) {
    if (Test-Path -LiteralPath $Path) {
        $null = $candidates.Add((Resolve-Path -LiteralPath $Path).Path)
        Say "    Указанная тобой папка принята: $Path"
    } else {
        Say "    ВНИМАНИЕ: указанная папка не найдена: $Path"
    }
}

# --- Способ 2: если TDM сейчас запущен, спросим у Windows его путь ---
Say "    Смотрю запущенные программы..."
try {
    $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match 'tdm' }
    foreach ($p in $procs) {
        try {
            if ($p.Path) {
                $dir = Split-Path -Parent $p.Path
                Say "    [+] TDM запущен: $($p.Path)"
                $null = $candidates.Add($dir)
            }
        } catch { }
    }
    if (-not $procs) { Say "    (запущенного TDM не вижу - это нормально, если он закрыт)" }
} catch {
    Say "    (не удалось прочитать список программ)"
}

# --- Способ 3: ярлык в меню "Пуск" ---
Say "    Смотрю ярлыки в меню Пуск..."
$startMenus = @(
    (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'),
    (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

foreach ($sm in $startMenus) {
    try {
        $lnks = Get-ChildItem -LiteralPath $sm -Recurse -Filter '*.lnk' -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match 'tdm' }
        foreach ($lnk in $lnks) {
            try {
                $shell = New-Object -ComObject WScript.Shell
                $target = $shell.CreateShortcut($lnk.FullName).TargetPath
                if ($target -and (Test-Path -LiteralPath $target)) {
                    Say "    [+] Ярлык '$($lnk.Name)' ведёт на: $target"
                    $null = $candidates.Add((Split-Path -Parent $target))
                }
            } catch { }
        }
    } catch { }
}

# --- Способ 4: поиск папок со словом tdm в обычных местах установки ---
Say "    Ищу папки со словом 'tdm' в местах установки программ..."
$roots = @(
    (Join-Path $env:LOCALAPPDATA 'Programs'),
    $env:LOCALAPPDATA,
    $env:APPDATA,
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)}
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

foreach ($root in $roots) {
    $hits = $null
    try {
        $hits = Get-ChildItem -LiteralPath $root -Directory -Recurse -Depth 3 -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match 'tdm' }
    } catch {
        # Запасной путь для очень старых версий PowerShell (без параметра -Depth)
        try {
            $hits = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match 'tdm' }
        } catch { }
    }
    foreach ($h in $hits) { $null = $candidates.Add($h.FullName) }
}

$found = $candidates | Select-Object -Unique

if (-not $found -or @($found).Count -eq 0) {
    Say ""
    Say "    Папку TDM автоматически найти не удалось."
    Say "    Сделай так: найди ярлык TDM, правой кнопкой -> 'Расположение файла',"
    Say "    скопируй адрес из строки проводника и запусти сканер так:"
    Say "      powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Path `"вставь_путь_сюда`""
    SaveReport
    return
}

Say ""
Say "    Возможные папки TDM:"
foreach ($f in $found) { Say "      - $f" }

$appDir = @($found)[0]
Say ""
Say "    Подробно осматриваю: $appDir"
Say ""

# ---------------------------------------------------------------------------
# ШАГ 2. Из чего собрана программа
# ---------------------------------------------------------------------------
Say "[2] Из чего собрано приложение"

$allFiles = @()
try {
    $allFiles = @(Get-ChildItem -LiteralPath $appDir -Recurse -File -ErrorAction SilentlyContinue)
} catch { }

Say "    Всего файлов внутри: $($allFiles.Count)"

# Признаки Electron: значит внутри программы обычный веб-код (JavaScript).
$electronMarkers = @('app.asar', 'electron.exe', 'ffmpeg.dll', 'libEGL.dll',
                     'chrome_100_percent.pak', 'LICENSES.chromium.html', 'v8_context_snapshot.bin')
$isElectron = $false
foreach ($m in $electronMarkers) {
    $hit = $allFiles | Where-Object { $_.Name -ieq $m } | Select-Object -First 1
    if ($hit) {
        $isElectron = $true
        Say "    [+] Признак Electron: $($hit.Name)  ->  $($hit.FullName)"
    }
}
if ($isElectron) {
    Say "    ВЫВОД: похоже, приложение сделано на Electron (веб-технологии)."
    Say "           Это хорошая новость: у таких приложений внутри бывает код,"
    Say "           который иногда даёт зацепки для автоматизации."
} else {
    Say "    ВЫВОД: явных признаков Electron не видно. Возможно, это обычная"
    Say "           программа (C++/C# и т.п.) - тогда зацепок меньше."
}

# Список исполняемых файлов - полезно понять состав программы
$exes = @($allFiles | Where-Object { $_.Extension -ieq '.exe' })
Say "    Программ (.exe) внутри: $($exes.Count)"
foreach ($e in ($exes | Select-Object -First 15)) {
    $mb = [math]::Round($e.Length / 1MB, 1)
    Say "      - $($e.Name)  [$mb МБ]"
}
Say ""

# ---------------------------------------------------------------------------
# ШАГ 3. Файлы настроек и упаковка кода
# ---------------------------------------------------------------------------
Say "[3] Файлы настроек и упаковки"

$pkg = $allFiles | Where-Object { $_.Name -ieq 'package.json' } | Select-Object -First 1
if ($pkg) {
    Say "    Найден package.json (паспорт Electron-приложения): $($pkg.FullName)"
    try {
        $lines = Get-Content -LiteralPath $pkg.FullName -ErrorAction SilentlyContinue
        Say "    --- содержимое (секреты скрыты) ---"
        foreach ($ln in ($lines | Select-Object -First 60)) { Say "      $(Mask $ln)" }
        Say "    --- конец ---"
    } catch { Say "    (не удалось прочитать)" }
} else {
    Say "    package.json не найден."
}

$asar = @($allFiles | Where-Object { $_.Extension -ieq '.asar' })
foreach ($a in $asar) {
    $mb = [math]::Round($a.Length / 1MB, 1)
    Say "    Упакованный код (.asar): $($a.FullName)  [$mb МБ]"
}

$configExt = @('.json', '.ini', '.cfg', '.conf', '.yaml', '.yml', '.xml', '.env', '.config', '.toml')
$configs = @($allFiles | Where-Object { $configExt -contains $_.Extension.ToLower() })
Say "    Файлов настроек найдено: $($configs.Count)"
foreach ($c in ($configs | Select-Object -First 40)) { Say "      - $($c.FullName)" }
if ($configs.Count -gt 40) { Say "      ... и ещё $($configs.Count - 40)" }
Say ""

# ---------------------------------------------------------------------------
# ШАГ 4. Ищем зацепки: bot / api / token / webhook / адреса
# ---------------------------------------------------------------------------
Say "[4] Поиск зацепок для подключения (bot / api / token / webhook / адреса)"
Say "    (возможные секреты скрыты звёздочками)"

$textExt = @('.json', '.js', '.ts', '.txt', '.ini', '.cfg', '.conf', '.yaml', '.yml',
             '.xml', '.env', '.config', '.toml', '.html', '.md', '.log')
$keywords = 'bot|api|token|webhook|endpoint|wss?://|https?://'

$scanFiles = @($allFiles | Where-Object {
    ($textExt -contains $_.Extension.ToLower()) -and ($_.Length -lt 5MB)
})
Say "    Просматриваю текстовых файлов: $($scanFiles.Count)"

$matchCount = 0
foreach ($file in $scanFiles) {
    if ($matchCount -ge 200) { break }
    try {
        $hits = Select-String -LiteralPath $file.FullName -Pattern $keywords -AllMatches -ErrorAction SilentlyContinue
        foreach ($h in $hits) {
            if ($matchCount -ge 200) { break }
            $line = "$($h.Line)".Trim()
            if ($line.Length -gt 200) { $line = $line.Substring(0, 200) + "..." }
            Say "      $($file.Name):$($h.LineNumber)  ->  $(Mask $line)"
            $matchCount++
        }
    } catch { }
}
if ($matchCount -ge 200) { Say "      ... (показаны первые 200 совпадений)" }
if ($matchCount -eq 0) {
    Say "      Ничего похожего в открытом виде не найдено."
    Say "      Если код упакован в .asar (см. шаг 3), зацепки могут быть внутри."
}
Say ""

Say "============================================================"
Say " ОСМОТР ЗАВЕРШЁН. Ничего не изменено, наружу ничего не отправлено."
Say " Пришли мне содержимое этого файла в чат."
Say "============================================================"

SaveReport
