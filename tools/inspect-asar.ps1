# ============================================================================
#  inspect-asar.ps1  -  БЕЗОПАСНЫЙ "открыватель коробки" app.asar
# ============================================================================
#
#  ЗАЧЕМ:
#    Весь код мессенджера TDM запечатан в один файл-архив app.asar.
#    Эта программа заглядывает внутрь и ищет там зацепки для подключения:
#    веб-адреса, "сокеты", слова bot / api / token / webhook.
#    Так мы поймём, есть ли у TDM вообще способ, к которому можно подключиться.
#
#  ЧЕГО НЕ ДЕЛАЕТ:
#    - НЕ распаковывает и НЕ меняет ни одного файла (только читает)
#    - НЕ запускает никакие программы
#    - НЕ отправляет НИЧЕГО в интернет
#    - НЕ собирает пароли: длинные "секретные" строки прячет звёздочками
#    - НЕ требует прав администратора
#
#  КАК ЗАПУСТИТЬ:
#    1) Открой "Windows PowerShell" (Пуск -> напечатать PowerShell).
#    2) Вставь команду и нажми Enter (путь - как в отчёте сканера):
#       powershell -ExecutionPolicy Bypass -File "C:\Users\ИМЯ_ПОЛЬЗОВАТЕЛЯ\Desktop\inspect-asar.ps1"
#
#    Если архив лежит в другом месте - укажи его прямо:
#       powershell -ExecutionPolicy Bypass -File "...\inspect-asar.ps1" -Asar "C:\Program Files\TDM\resources\app.asar"
#
#  РЕЗУЛЬТАТ: на Рабочем столе появится TDM_asar_report.txt.
#  Открой его, посмотри глазами (внутри могут быть адреса серверов),
#  и пришли мне в чат.
# ============================================================================

param(
    [string]$Asar = "C:\Program Files\TDM\resources\app.asar"
)

$ErrorActionPreference = 'Stop'

$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrEmpty($desktop)) { $desktop = Join-Path $env:USERPROFILE 'Desktop' }
$reportPath = Join-Path $desktop 'TDM_asar_report.txt'

$report = New-Object System.Collections.Generic.List[string]
function Say([string]$line) { Write-Host $line; $null = $report.Add($line) }
function SaveReport() {
    try { $report | Out-File -FilePath $reportPath -Encoding UTF8; Write-Host ""; Write-Host "Отчёт сохранён: $reportPath" }
    catch { Write-Host "Не удалось сохранить отчёт: $($_.Exception.Message)" }
}
# Прячем длинные "секретные" строки (32+ символа подряд).
function Mask([string]$t) {
    if ([string]::IsNullOrEmpty($t)) { return "" }
    return [System.Text.RegularExpressions.Regex]::Replace($t, '[A-Za-z0-9_\+/=]{32,}', '***СКРЫТО***')
}

Say "============================================================"
Say " ОТКРЫВАТЕЛЬ КОРОБКИ app.asar - отчёт"
Say " Дата: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Say " Система: $([Environment]::OSVersion.VersionString)"
Say " PowerShell: $($PSVersionTable.PSVersion)"
Say "============================================================"
Say ""

if (-not (Test-Path -LiteralPath $Asar)) {
    Say "ОШИБКА: не найден файл архива: $Asar"
    Say "Проверь путь. В отчёте сканера он был в разделе 'Упакованный код (.asar)'."
    SaveReport; return
}

$fileInfo = Get-Item -LiteralPath $Asar
$sizeMB = [math]::Round($fileInfo.Length / 1MB, 1)
Say "Архив: $Asar"
Say "Размер: $sizeMB МБ"
Say ""

# ---------------------------------------------------------------------------
# ЧАСТЬ 1. Читаем "оглавление" архива (список файлов внутри)
# ---------------------------------------------------------------------------
Say "[1] Оглавление архива (что за файлы лежат внутри)"

$fs = $null
$jsonOk = $false
$baseOffset = 0
$tree = $null
try {
    $fs = [System.IO.File]::OpenRead($Asar)
    $head = New-Object byte[] 16
    $null = $fs.Read($head, 0, 16)
    # Формат asar (два вложенных "пикла"):
    #   [0:4]  = 4          (служебное)
    #   [4:8]  = S          (размер оглавления в байтах)
    #   [8:12] = S-4        (служебное)
    #   [12:16]= длина JSON-строки
    #   с [16:] начинается сам JSON-текст оглавления
    $S       = [BitConverter]::ToUInt32($head, 4)
    $jsonLen = [BitConverter]::ToUInt32($head, 12)
    if ($jsonLen -gt 0 -and $jsonLen -lt $fileInfo.Length) {
        $jsonBytes = New-Object byte[] $jsonLen
        # позиция уже на 16 после чтения заголовка - читаем JSON подряд
        $read = 0
        while ($read -lt $jsonLen) {
            $n = $fs.Read($jsonBytes, $read, $jsonLen - $read)
            if ($n -le 0) { break }
            $read += $n
        }
        $jsonText = [System.Text.Encoding]::UTF8.GetString($jsonBytes, 0, $read)
        $tree = $jsonText | ConvertFrom-Json
        # Данные начинаются после size-пикла (8 б.) и оглавления (S б.)
        $baseOffset = 8 + $S
        $jsonOk = $true
    }
} catch {
    Say "    Не удалось прочитать оглавление: $($_.Exception.Message)"
}

# Разворачиваем дерево файлов в простой список "путь|размер|смещение"
$flat = New-Object System.Collections.Generic.List[object]
function Walk($node, [string]$prefix) {
    if ($null -eq $node) { return }
    $filesProp = $node.PSObject.Properties['files']
    if ($null -eq $filesProp -or $null -eq $filesProp.Value) { return }
    foreach ($p in $filesProp.Value.PSObject.Properties) {
        $child = $p.Value
        $path = if ($prefix -eq "") { $p.Name } else { "$prefix/$($p.Name)" }
        $hasFiles = $child.PSObject.Properties['files']
        if ($hasFiles) {
            Walk $child $path
        } else {
            # Берём size/offset через отражение (надёжнее, чем $child.size)
            $sz = 0; $off = -1
            $szProp  = $child.PSObject.Properties['size']
            $offProp = $child.PSObject.Properties['offset']
            if ($szProp)  { $sz  = [int64]$szProp.Value }
            if ($offProp) { $off = [int64]$offProp.Value }
            $null = $flat.Add([pscustomobject]@{ Path = $path; Size = $sz; Offset = $off })
        }
    }
}

if ($jsonOk) {
    try { Walk $tree "" } catch { Say "    (не удалось разобрать оглавление полностью)" }
    Say "    Файлов внутри архива: $($flat.Count)"

    # Показываем "интересные" файлы: настройки и то, что может касаться связи
    $interesting = $flat | Where-Object {
        $_.Path -match '(?i)(config|setting|\.env|\.json$|\.yml$|\.yaml$|api|socket|websocket|\bws\b|bot|webhook|preload|main\.js$)' -and
        $_.Path -notmatch '(?i)node_modules'
    }
    Say "    Похоже на настройки/связь (без служебных библиотек): $($interesting.Count)"
    foreach ($f in ($interesting | Select-Object -First 60)) {
        Say ("      - {0}  [{1} байт]" -f $f.Path, $f.Size)
    }
    if ($interesting.Count -gt 60) { Say "      ... и ещё $($interesting.Count - 60)" }
} else {
    Say "    Оглавление прочитать не вышло - перейду сразу к поиску по тексту (часть 3)."
}
Say ""

# Вспомогательная: прочитать один небольшой файл из архива по смещению
function Read-AsarText([int64]$offset, [int64]$size) {
    if ($offset -lt 0 -or $size -le 0 -or $size -gt 2MB) { return $null }
    try {
        $buf = New-Object byte[] $size
        $fs.Position = $baseOffset + $offset
        $read = 0
        while ($read -lt $size) {
            $n = $fs.Read($buf, $read, $size - $read)
            if ($n -le 0) { break }
            $read += $n
        }
        return [System.Text.Encoding]::UTF8.GetString($buf, 0, $read)
    } catch { return $null }
}

# ---------------------------------------------------------------------------
# ЧАСТЬ 2. Читаем "паспорт" самого приложения (его package.json)
# ---------------------------------------------------------------------------
Say "[2] Паспорт приложения (package.json самого TDM, не служебных библиотек)"
if ($jsonOk) {
    $ownPkg = $flat | Where-Object {
        ($_.Path -eq 'package.json' -or $_.Path -eq 'app/package.json') -and $_.Path -notmatch 'node_modules'
    } | Select-Object -First 1
    if ($ownPkg) {
        $txt = Read-AsarText $ownPkg.Offset $ownPkg.Size
        if ($txt) {
            Say "    Файл: $($ownPkg.Path)"
            Say "    --- содержимое (секреты скрыты) ---"
            foreach ($ln in (($txt -split "`n") | Select-Object -First 80)) { Say "      $(Mask $ln.TrimEnd())" }
            Say "    --- конец ---"
        } else { Say "    (нашёл, но не смог прочитать)" }
    } else {
        Say "    Собственный package.json в оглавлении не найден."
    }
} else {
    Say "    Пропущено (оглавление не прочиталось)."
}
Say ""

# ---------------------------------------------------------------------------
# ЧАСТЬ 3. Поиск зацепок по всему тексту архива
# ---------------------------------------------------------------------------
Say "[3] Поиск зацепок по всему архиву (адреса, сокеты, bot/api/token/webhook)"
Say "    (возможные секреты скрыты; читаю кусками, это может занять минуту)"

$urlRegex = New-Object System.Text.RegularExpressions.Regex '(?:https?|wss?)://[^\s"''<>\\\)\]]{2,200}'
# Ключевые слова: "bot" считаем только как отдельное слово, чтобы не ловить
# bottom/botany/botswana и прочий мусор из библиотек.
$kwRegex  = New-Object System.Text.RegularExpressions.Regex '(?i)(socket\.io|websocket|webhook|/api/|Bearer\b|Authorization|oauth|access[_-]?token|client[_-]?secret|\bbots?\b|bot[_-]?api|bot[_-]?token|chat\s?bot)'

# Что считаем "своим/интересным" адресом (а не документацией библиотек)
$hotRegex   = New-Object System.Text.RegularExpressions.Regex '(?i)(mos\.ru|\btdm|/api/|wss?://|127\.0\.0\.1|localhost)'
# Явный "шум" - ссылки на документацию открытых библиотек
$boringRegex = New-Object System.Text.RegularExpressions.Regex '(?i)(github|wikipedia|ecma-international|w3\.org|mozilla|python\.org|amazon|aws|google|goo\.gl|bower\.io|brew\.sh|nodejs|npmjs|jquery|stackoverflow|creativecommons|editorconfig|fsf\.org|feross|izs\.me|codeclimate|testling|fb\.me|nist\.gov|garant\.ru|example\.(com|org)|foo\.com|\.spec\.whatwg)'

$urls = New-Object System.Collections.Generic.HashSet[string]   # все адреса
$hot  = New-Object System.Collections.Generic.HashSet[string]   # свои/интересные
$kwCounts = @{}

try {
    if ($null -eq $fs) { $fs = [System.IO.File]::OpenRead($Asar) }
    $fs.Position = 0
    $chunkSize = 4MB
    $buffer = New-Object byte[] $chunkSize
    $tail = ""
    $done = 0
    while ($true) {
        $n = $fs.Read($buffer, 0, $chunkSize)
        if ($n -le 0) { break }
        # Превращаем байты в текст (нечитаемые символы станут мусором - это ок)
        $text = $tail + [System.Text.Encoding]::ASCII.GetString($buffer, 0, $n)

        foreach ($m in $urlRegex.Matches($text)) {
            $u = $m.Value.TrimEnd('.',',',';')
            if ($u.Length -le 200) {
                $mu = Mask $u
                $null = $urls.Add($mu)
                if ($hotRegex.IsMatch($u)) { $null = $hot.Add($mu) }
            }
        }
        foreach ($m in $kwRegex.Matches($text)) {
            $k = $m.Value.ToLower().Trim()
            if ($kwCounts.ContainsKey($k)) { $kwCounts[$k]++ } else { $kwCounts[$k] = 1 }
        }

        # Хвост, чтобы не потерять адрес на стыке кусков
        if ($text.Length -gt 300) { $tail = $text.Substring($text.Length - 300) } else { $tail = $text }

        $done += $n
        if ($done % (40MB) -lt $chunkSize) {
            Say "    ...просмотрено $([math]::Round($done/1MB)) МБ из $sizeMB МБ"
        }
    }
} catch {
    Say "    Ошибка при чтении: $($_.Exception.Message)"
} finally {
    if ($fs) { $fs.Close() }
}

Say ""
Say "    НАЙДЕННЫЕ КЛЮЧЕВЫЕ СЛОВА (сколько раз встретились):"
if ($kwCounts.Count -eq 0) {
    Say "      (ничего из списка не найдено)"
} else {
    foreach ($k in ($kwCounts.Keys | Sort-Object { -$kwCounts[$_] })) {
        Say ("      {0,-20} {1}" -f $k, $kwCounts[$k])
    }
}
Say ""
Say "    >>> ГЛАВНОЕ: СВОИ АДРЕСА TDM (сервера, api, веб-сокеты, localhost):"
if ($hot.Count -eq 0) {
    Say "      (не найдено - возможно, адреса собираются в коде из кусков)"
} else {
    foreach ($u in ($hot | Sort-Object | Select-Object -First 200)) { Say "      $u" }
    if ($hot.Count -gt 200) { Say "      ... (показаны первые 200 из $($hot.Count))" }
}
Say ""
Say "    ПРОЧИЕ НЕОБЫЧНЫЕ АДРЕСА (без явной документации библиотек, до 80):"
$other = $urls | Where-Object { -not $hot.Contains($_) -and -not $boringRegex.IsMatch($_) }
if (-not $other -or @($other).Count -eq 0) {
    Say "      (только ссылки на документацию библиотек - не показываю)"
} else {
    foreach ($u in (@($other) | Sort-Object | Select-Object -First 80)) { Say "      $u" }
    if (@($other).Count -gt 80) { Say "      ... (показаны первые 80 из $(@($other).Count))" }
}
Say ""
Say "    Всего уникальных адресов найдено: $($urls.Count) (из них своих/интересных: $($hot.Count))"
Say ""

Say "============================================================"
Say " ГОТОВО. Ничего не изменено, наружу ничего не отправлено."
Say " Пришли мне содержимое этого файла в чат."
Say "============================================================"

SaveReport
