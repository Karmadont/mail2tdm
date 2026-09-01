# ============================================================================
#  scan-tdm.ps1  —  БЕЗОПАСНЫЙ "осмотрщик" приложения TDM Messenger
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
#    - НЕ собирает ваши пароли: длинные "секретные" строки он прячет звёздочками
#    - НЕ требует прав администратора
#
#  КАК ЗАПУСТИТЬ (два шага):
#    1) Нажми "Пуск", напечатай  PowerShell , открой "Windows PowerShell".
#    2) Перетащи в открывшееся окно этот файл и нажми Enter.
#       (или напечатай:  powershell -ExecutionPolicy Bypass -File "путь\к\scan-tdm.ps1" )
#
#    Если TDM установлен в необычном месте и сканер его не нашёл — укажи папку сам:
#       powershell -ExecutionPolicy Bypass -File "scan-tdm.ps1" -Path "C:\путь\к\папке TDM"
#
#  КОГДА ЗАКОНЧИТ: открой файл  TDM_scan_report.txt  на Рабочем столе,
#  скопируй его содержимое и пришли мне в чат.
# ============================================================================

param(
    # Необязательно: прямой путь к папке TDM, если сканер её не нашёл сам
    [string]$Path = ""
)

# --- Куда сохраняем отчёт: на Рабочий стол, чтобы легко найти ---
$reportPath = Join-Path $env:USERPROFILE "Desktop\TDM_scan_report.txt"
$report = New-Object System.Collections.Generic.List[string]

function Say([string]$line) {
    # Пишем строку и на экран, и в отчёт
    Write-Host $line
    $report.Add($line)
}

# Прячем длинные "секретные" строки (похожие на токены/ключи), чтобы пароли
# не попали в отчёт. Заменяем цепочки из 20+ букв/цифр на звёздочки.
function Mask([string]$text) {
    if ($null -eq $text) { return "" }
    return [System.Text.RegularExpressions.Regex]::Replace(
        $text, '[A-Za-z0-9_\-\.]{20,}', '***СКРЫТО***')
}

Say "============================================================"
Say " ОСМОТР TDM MESSENGER — отчёт"
Say " Дата: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Say " Компьютер: $env:COMPUTERNAME   Пользователь: $env:USERNAME"
Say "============================================================"
Say ""

# ---------------------------------------------------------------------------
# ШАГ 1. Ищем папку приложения TDM
# ---------------------------------------------------------------------------
Say "[1] Ищу, где установлен TDM..."

$found = @()

if ($Path -ne "") {
    if (Test-Path $Path) {
        $found = @($Path)
        Say "    Использую папку, которую ты указал: $Path"
    } else {
        Say "    ВНИМАНИЕ: папка не найдена: $Path"
    }
}

if ($found.Count -eq 0) {
    # Обычные места, куда программы ставятся БЕЗ прав администратора и с ними
    $roots = @(
        (Join-Path $env:LOCALAPPDATA "Programs"),
        $env:LOCALAPPDATA,
        $env:APPDATA,
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)}
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    foreach ($root in $roots) {
        # Ищем папки, в названии которых есть "tdm" (не глубже 3 уровней — чтобы быстро)
        try {
            $hits = Get-ChildItem -Path $root -Directory -Recurse -Depth 3 -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match 'tdm' }
            foreach ($h in $hits) { $found += $h.FullName }
        } catch {}
    }
    $found = $found | Select-Object -Unique
}

if ($found.Count -eq 0) {
    Say ""
    Say "    Папку TDM автоматически найти не удалось."
    Say "    Подскажи путь вручную. Как узнать путь:"
    Say "      - найди ярлык TDM на Рабочем столе,"
    Say "      - правой кнопкой -> 'Расположение файла',"
    Say "      - скопируй адрес из строки проводника,"
    Say "      - запусти сканер так:"
    Say "        powershell -ExecutionPolicy Bypass -File `"scan-tdm.ps1`" -Path `"вставь_путь_сюда`""
    Say ""
    Say "Готово. Отчёт сохранён: $reportPath"
    $report | Out-File -FilePath $reportPath -Encoding UTF8
    return
}

Say "    Нашёл возможные папки TDM:"
foreach ($f in $found) { Say "      - $f" }
# Берём первую найденную как основную
$appDir = $found[0]
Say ""
Say "    Осматриваю: $appDir"
Say ""

# ---------------------------------------------------------------------------
# ШАГ 2. Из чего собрана программа? (Electron / обычное приложение)
# ---------------------------------------------------------------------------
Say "[2] Из чего собрано приложение"

$allFiles = @()
try {
    $allFiles = Get-ChildItem -Path $appDir -Recurse -File -ErrorAction SilentlyContinue
} catch {}

Say "    Всего файлов внутри: $($allFiles.Count)"

# Признаки Electron (это значит, что внутри программы — обычный веб-код на
# JavaScript; такие приложения иногда можно расширять). Ищем характерные файлы.
$electronMarkers = @('app.asar','electron.exe','ffmpeg.dll','libEGL.dll',
                     'chrome_100_percent.pak','LICENSES.chromium.html','v8_context_snapshot.bin')
$isElectron = $false
foreach ($m in $electronMarkers) {
    $hit = $allFiles | Where-Object { $_.Name -ieq $m } | Select-Object -First 1
    if ($hit) {
        $isElectron = $true
        Say "    [+] Найден признак Electron: $($hit.Name)  ->  $($hit.FullName)"
    }
}
if ($isElectron) {
    Say "    ВЫВОД: похоже, приложение сделано на Electron (веб-технологии)."
    Say "           Это хорошая новость: у таких приложений внутри бывает код,"
    Say "           который иногда даёт зацепки для автоматизации."
} else {
    Say "    ВЫВОД: явных признаков Electron не видно. Возможно, это обычная"
    Say "           программа (написана на C++/C# и т.п.) — тогда зацепок меньше."
}
Say ""

# ---------------------------------------------------------------------------
# ШАГ 3. Файлы настроек и упаковка кода
# ---------------------------------------------------------------------------
Say "[3] Файлы настроек и упаковки"

# package.json — "паспорт" Electron-приложения, там бывает имя и версия
$pkg = $allFiles | Where-Object { $_.Name -ieq 'package.json' } | Select-Object -First 1
if ($pkg) {
    Say "    Найден package.json: $($pkg.FullName)"
    try {
        $content = Get-Content $pkg.FullName -Raw -ErrorAction SilentlyContinue
        Say "    --- содержимое package.json (секреты скрыты) ---"
        foreach ($ln in ($content -split "`n")) { Say "      $(Mask $ln)" }
        Say "    --- конец package.json ---"
    } catch { Say "    (не удалось прочитать)" }
}

# asar — "коробка", в которую упакован код приложения
$asar = $allFiles | Where-Object { $_.Extension -ieq '.asar' }
foreach ($a in $asar) {
    $sizeMB = [math]::Round($a.Length / 1MB, 1)
    Say "    Упакованный код (.asar): $($a.FullName)  [$sizeMB МБ]"
}

# Файлы настроек
$configExt = @('.json','.ini','.cfg','.conf','.yaml','.yml','.xml','.env','.config','.toml')
$configs = $allFiles | Where-Object { $configExt -contains $_.Extension.ToLower() }
Say "    Файлов настроек найдено: $($configs.Count)"
foreach ($c in ($configs | Select-Object -First 40)) {
    Say "      - $($c.FullName)"
}
if ($configs.Count -gt 40) { Say "      ... и ещё $($configs.Count - 40)" }
Say ""

# ---------------------------------------------------------------------------
# ШАГ 4. Ищем зацепки: bot / api / token / webhook / веб-адреса
# ---------------------------------------------------------------------------
Say "[4] Поиск зацепок для подключения (bot / api / token / webhook / адреса)"
Say "    (значения возможных секретов скрыты звёздочками)"

# Ищем только в текстовых файлах разумного размера (до 5 МБ), чтобы не читать
# тяжёлые двоичные файлы и не тормозить.
$textExt = @('.json','.js','.ts','.txt','.ini','.cfg','.conf','.yaml','.yml',
             '.xml','.env','.config','.toml','.html','.md','.log')
$keywords = 'bot|api|token|webhook|endpoint|wss?://|https?://'

$scanFiles = $allFiles | Where-Object {
    ($textExt -contains $_.Extension.ToLower()) -and ($_.Length -lt 5MB)
}
Say "    Просматриваю текстовых файлов: $($scanFiles.Count)"

$matchCount = 0
foreach ($file in $scanFiles) {
    try {
        $hits = Select-String -Path $file.FullName -Pattern $keywords -AllMatches -ErrorAction SilentlyContinue
        foreach ($h in $hits) {
            if ($matchCount -ge 200) { break }   # не раздуваем отчёт
            $line = ($h.Line).Trim()
            if ($line.Length -gt 200) { $line = $line.Substring(0,200) + "..." }
            Say "      $($file.Name):$($h.LineNumber)  ->  $(Mask $line)"
            $matchCount++
        }
    } catch {}
    if ($matchCount -ge 200) { Say "      ... (показаны первые 200 совпадений)"; break }
}
if ($matchCount -eq 0) {
    Say "      Ничего похожего не найдено в открытом виде."
    Say "      Если код упакован в .asar (см. шаг 3), зацепки могут быть внутри коробки —"
    Say "      напиши мне, и я подскажу безопасный способ заглянуть и туда."
}
Say ""

# ---------------------------------------------------------------------------
# Итог
# ---------------------------------------------------------------------------
Say "============================================================"
Say " ОСМОТР ЗАВЕРШЁН. Ничего не изменено, наружу ничего не отправлено."
Say " Отчёт сохранён: $reportPath"
Say " Пришли мне содержимое этого файла в чат."
Say "============================================================"

$report | Out-File -FilePath $reportPath -Encoding UTF8
