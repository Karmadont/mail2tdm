# ============================================================================
#  tdm-check.ps1  -  проверка связи с TDM Bot API
# ============================================================================
#
#  ЗАЧЕМ:
#    Когда техподдержка выдаст доступ и бот Bot Creator выдаст токен -
#    запусти это ПЕРВЫМ делом. Программа проверит:
#      1) читается ли файл настроек;
#      2) принимает ли сервер твой токен;
#      3) в каких группах/каналах состоит бот (и покажет их ID).
#    И по желанию отправит одно тестовое сообщение в канал.
#
#  ЧЕГО НЕ ДЕЛАЕТ:
#    - НЕ показывает твой токен на экране (прячет звёздочками)
#    - НЕ пишет токен в журнал
#    - Без ключа -Send ничего никуда не отправляет
#
#  КАК ЗАПУСТИТЬ (только проверка, без отправки):
#    powershell -ExecutionPolicy Bypass -File "C:\Users\ИМЯ_ПОЛЬЗОВАТЕЛЯ\Desktop\tdm-check.ps1"
#
#  С отправкой тестового сообщения в канал:
#    powershell -ExecutionPolicy Bypass -File "...\tdm-check.ps1" -Send
#
#  Настройки берутся из  %APPDATA%\mail2tdm\config.ini
# ============================================================================

param(
    # Отправить тестовое сообщение в канал (без этого ключа - только проверка)
    [switch]$Send,
    # Свой путь к файлу настроек (по умолчанию %APPDATA%\mail2tdm\config.ini)
    [string]$Config = ""
)

# Современное шифрование соединения: без этого старый Windows может не достучаться
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch { }

function Say([string]$t) { Write-Host $t }

# --- 1. Читаем настройки -----------------------------------------------
if ([string]::IsNullOrEmpty($Config)) {
    $Config = Join-Path $env:APPDATA 'mail2tdm\config.ini'
}

Say "============================================================"
Say " ПРОВЕРКА СВЯЗИ С TDM BOT API"
Say " Файл настроек: $Config"
Say "============================================================"
Say ""

if (-not (Test-Path -LiteralPath $Config)) {
    Say "ОШИБКА: файл настроек не найден."
    Say "Создай его по образцу config.example.ini и заполни."
    Say "Как попасть в папку: Win+R -> вставь  %APPDATA%  -> Enter -> создай папку mail2tdm"
    return
}

# Разбираем строки вида Имя=Значение (строки с ; и # - комментарии)
$cfg = @{}
foreach ($line in (Get-Content -LiteralPath $Config -Encoding UTF8)) {
    $ln = $line.Trim()
    if ($ln.Length -eq 0) { continue }
    if ($ln.StartsWith(';') -or $ln.StartsWith('#')) { continue }
    $p = $ln.IndexOf('=')
    if ($p -gt 0) {
        $k = $ln.Substring(0, $p).Trim()
        $v = $ln.Substring($p + 1).Trim()
        $cfg[$k] = $v
    }
}

function Cfg([string]$name) { if ($cfg.ContainsKey($name)) { return $cfg[$name] } else { return "" } }

$baseUrl     = (Cfg 'BaseUrl').TrimEnd('/')
$token       = Cfg 'Token'
$workspaceId = Cfg 'WorkspaceId'
$groupId     = Cfg 'GroupId'

# Показываем, что прочитали. Токен НЕ показываем.
Say "[1] Что прочитано из настроек"
Say "    Адрес сервера : $(if($baseUrl){$baseUrl}else{'(ПУСТО)'})"
if ([string]::IsNullOrEmpty($token)) {
    Say "    Токен         : (ПУСТО)"
} else {
    Say "    Токен         : задан, длина $($token.Length) символов (значение скрыто)"
}
Say "    ID пространства: $(if($workspaceId){$workspaceId}else{'(ПУСТО)'})"
Say "    ID группы      : $(if($groupId){$groupId}else{'(ПУСТО)'})"
Say ""

$missing = @()
if (-not $baseUrl) { $missing += 'BaseUrl' }
if (-not $token)   { $missing += 'Token' }
if ($missing.Count -gt 0) {
    Say "ОШИБКА: не заполнено: $($missing -join ', ')"
    Say "Заполни эти поля в файле настроек и запусти снова."
    return
}

$headers = @{ 'Authorization' = $token }

# --- 2. Спрашиваем у сервера список групп бота -------------------------
Say "[2] Спрашиваю сервер, в каких группах состоит бот"
$groupsUrl = "$baseUrl/botapi/v1/groups/getAllUserGroupStates"
Say "    Запрос: POST $groupsUrl"

$states = $null
try {
    $states = Invoke-RestMethod -Uri $groupsUrl -Method Post -Headers $headers `
                                -ContentType 'application/json; charset=utf-8' -TimeoutSec 30
    Say "    Сервер ответил успешно."
} catch {
    Say ""
    Say "    НЕ ПОЛУЧИЛОСЬ: $($_.Exception.Message)"
    $resp = $_.Exception.Response
    if ($resp) {
        try {
            $code = [int]$resp.StatusCode
            Say "    Код ответа: $code"
            if ($code -eq 401 -or $code -eq 403) {
                Say "    Похоже, сервер не принял токен. Проверь, что скопировал его целиком."
            }
        } catch { }
    } else {
        Say "    Сервер вообще не ответил. Проверь адрес BaseUrl и доступ к сети с рабочего ПК."
    }
    return
}

$list = @($states)
Say "    Групп найдено: $($list.Count)"
if ($list.Count -gt 0) {
    Say ""
    Say "    ID групп, в которых состоит бот (пригодятся для настройки GroupId):"
    $n = 0
    foreach ($g in $list) {
        $gid = $null
        $p = $g.PSObject.Properties['groupId']
        if ($p) { $gid = $p.Value }
        if ($gid) { Say "      - $gid"; $n++ }
        if ($n -ge 50) { Say "      ... (показаны первые 50)"; break }
    }
} else {
    Say ""
    Say "    Бот пока никуда не добавлен. Добавь его в нужный канал в TDM."
}
Say ""

# --- 3. По желанию - тестовое сообщение --------------------------------
if (-not $Send) {
    Say "[3] Тестовое сообщение НЕ отправлялось."
    Say "    Чтобы отправить, добавь ключ -Send к команде запуска."
    Say ""
    Say "ПРОВЕРКА ЗАВЕРШЕНА."
    return
}

Say "[3] Отправляю тестовое сообщение в канал"
if (-not $workspaceId -or -not $groupId) {
    Say "    НЕ МОГУ: не заполнены WorkspaceId и/или GroupId в настройках."
    Say "    Как узнать: открой веб-версию TDM, зайди в нужный канал -"
    Say "    оба числа будут прямо в адресной строке браузера."
    return
}

$sendUrl = "$baseUrl/botapi/v1/messages/sendTextMessage/$workspaceId/$groupId"
$payload = @{ message = 'Проверка связи: mail2tdm подключён.' } | ConvertTo-Json -Compress
# Отправляем байтами в UTF-8, иначе русский текст придёт испорченным
$bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)

Say "    Запрос: POST $sendUrl"
try {
    $r = Invoke-RestMethod -Uri $sendUrl -Method Post -Headers $headers `
                           -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec 30
    $mid = $null
    if ($r) { $p = $r.PSObject.Properties['messageId']; if ($p) { $mid = $p.Value } }
    if ($mid) {
        Say "    ОТПРАВЛЕНО. Сервер присвоил сообщению номер: $mid"
    } else {
        Say "    Сервер принял запрос. Ответ: $($r | ConvertTo-Json -Compress)"
    }
    Say ""
    Say "    Загляни в канал TDM - сообщение должно быть там."
} catch {
    Say "    НЕ ПОЛУЧИЛОСЬ: $($_.Exception.Message)"
    $resp = $_.Exception.Response
    if ($resp) { try { Say "    Код ответа: $([int]$resp.StatusCode)" } catch { } }
    Say "    Проверь, что бот добавлен в этот канал и что ID указаны верно."
}
Say ""
Say "ПРОВЕРКА ЗАВЕРШЕНА."
