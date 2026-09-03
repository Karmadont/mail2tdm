"""
mail2tdm.py - пересылка писем из Outlook в канал мессенджера TDM.

КАК ЗАПУСКАТЬ (открой "Командную строку" и набери одну из команд):

  python mail2tdm.py --selftest
        Самопроверка. НЕ нужны ни токен, ни интернет, ни Outlook.
        Проверяет, что программа исправна. Запусти это первым делом.

  python mail2tdm.py --check
        Проверка связи с TDM: принимает ли сервер токен и в каких
        каналах состоит бот. Ничего никуда не отправляет.

  python mail2tdm.py --test-message
        Отправить одно тестовое сообщение в канал.

  python mail2tdm.py --once
        Разобрать новые письма один раз и завершиться.

  python mail2tdm.py --watch
        Работать постоянно: проверять почту каждые N секунд.

БЕЗОПАСНОСТЬ:
  Пока в настройках TestMode=1, программа НИЧЕГО НЕ ОТПРАВЛЯЕТ - только
  пишет в журнал, что было бы отправлено. Это защита от случайной рассылки.

Настройки берутся из файла config.ini рядом с программой,
либо из %APPDATA%\\mail2tdm\\config.ini
"""

import argparse
import json
import os
import sys
import time
import traceback
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tdm_client import TdmClient, TdmError  # noqa: E402


# =====================================================================
#  Настройки и журнал
# =====================================================================

def find_config(explicit=None):
    """Ищем файл настроек: указанный -> рядом с программой -> в %APPDATA%."""
    if explicit:
        return explicit
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
    if os.path.isfile(here):
        return here
    appdata = os.environ.get("APPDATA", "")
    if appdata:
        return os.path.join(appdata, "mail2tdm", "config.ini")
    return here


def read_config(path):
    """Читаем строки вида Имя=Значение. Строки с ; и # - комментарии."""
    cfg = {}
    if not os.path.isfile(path):
        return cfg
    with open(path, "r", encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line[0] in ";#[":
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                cfg[k.strip()] = v.strip()
    return cfg


class Logger:
    """Пишет и на экран, и в файл журнала."""

    def __init__(self, path):
        self.path = path
        d = os.path.dirname(path)
        if d and not os.path.isdir(d):
            try:
                os.makedirs(d, exist_ok=True)
            except OSError:
                pass

    def __call__(self, msg):
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = "%s  %s" % (stamp, msg)
        print(line, flush=True)
        try:
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass


# =====================================================================
#  Чтение писем из Outlook
# =====================================================================

def get_outlook_folder(cfg, log):
    """
    Подключаемся к Outlook, который уже запущен и в котором вы уже вошли.
    Пароли не нужны: используется ваш открытый профиль Outlook.
    """
    try:
        import win32com.client  # входит в пакет pywin32
    except ImportError:
        raise TdmError(
            "Не установлена надстройка для чтения Outlook.\n"
            "Установи её командой:   pip install --user pywin32\n"
            "(права администратора не нужны)")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    mailbox = cfg.get("MailboxName", "").strip()
    folder_path = cfg.get("FolderPath", "").strip()

    if not mailbox:
        folder = outlook.GetDefaultFolder(6)  # 6 = Входящие
    else:
        folder = None
        for i in range(1, outlook.Stores.Count + 1):
            store = outlook.Stores.Item(i)
            if mailbox.lower() in str(store.DisplayName).lower():
                folder = store.GetDefaultFolder(6)
                break
        if folder is None:
            log("НЕ НАЙДЕН ящик '%s' - беру основной ящик." % mailbox)
            folder = outlook.GetDefaultFolder(6)

    if folder_path:
        for part in folder_path.split("\\"):
            part = part.strip()
            if part:
                folder = folder.Folders(part)

    return folder


def format_message(mail, cfg):
    """Собираем текст сообщения для канала."""
    def safe(getter, default=""):
        try:
            value = getter()
            return default if value is None else str(value)
        except Exception:
            return default

    subject = safe(lambda: mail.Subject, "(без темы)")
    sender = safe(lambda: mail.SenderEmailAddress) or safe(lambda: mail.SenderName)
    try:
        received = mail.ReceivedTime.strftime("%d.%m.%Y %H:%M")
    except Exception:
        received = datetime.now().strftime("%d.%m.%Y %H:%M")
    body = safe(lambda: mail.Body)

    # схлопываем лишние пустые строки
    while "\n\n\n" in body:
        body = body.replace("\n\n\n", "\n\n")
    body = body.strip()

    parts = [
        "Новая задача из почты",
        "Тема: %s" % subject,
        "От: %s" % sender,
        "Получено: %s" % received,
    ]

    names = attachment_names(mail)
    if names:
        parts.append("Вложения (%d): %s" % (len(names), ", ".join(names)))

    text = "\n".join(parts) + "\n\n" + body

    max_len = int(cfg.get("MaxLength", "3000") or 3000)
    if max_len > 0 and len(text) > max_len:
        text = text[:max_len - 20] + "\n... (текст обрезан)"
    return text


def attachment_names(mail):
    names = []
    try:
        count = mail.Attachments.Count
        for i in range(1, count + 1):
            names.append(str(mail.Attachments.Item(i).FileName))
    except Exception:
        pass
    return names


def save_attachments(mail, folder, log):
    """Сохраняем вложения во временную папку. Возвращаем список путей."""
    saved = []
    try:
        os.makedirs(folder, exist_ok=True)
        count = mail.Attachments.Count
        for i in range(1, count + 1):
            att = mail.Attachments.Item(i)
            name = str(att.FileName)
            # убираем опасные символы из имени файла
            safe_name = "".join(c for c in name if c not in '\\/:*?"<>|')
            path = os.path.join(folder, "%d_%s" % (int(time.time() * 1000) % 100000, safe_name))
            att.SaveAsFile(path)
            saved.append(path)
    except Exception as e:
        log("Не удалось сохранить вложения: %s" % e)
    return saved


# =====================================================================
#  Память об обработанных письмах (чтобы не отправить дважды)
# =====================================================================

class SeenStore:
    def __init__(self, path, log):
        self.path = path
        self.log = log
        self.seen = set()
        if os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    self.seen = set(json.load(f))
            except Exception:
                self.seen = set()

    def has(self, key):
        return key in self.seen

    def add(self, key):
        self.seen.add(key)
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(sorted(self.seen)[-5000:], f)
        except OSError as e:
            self.log("Не удалось сохранить список обработанных писем: %s" % e)


# =====================================================================
#  Основная работа
# =====================================================================

def process_once(cfg, client, log, seen, test_mode):
    folder = get_outlook_folder(cfg, log)
    log("Смотрю папку: %s" % folder.FolderPath)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)   # новые сверху

    limit = int(cfg.get("BatchLimit", "20") or 20)
    send_files = cfg.get("SendAttachments", "0").strip() == "1"
    tmp_dir = cfg.get("TempDir", "").strip() or os.path.join(
        os.environ.get("TEMP", "."), "mail2tdm")

    handled = 0
    for index in range(1, min(items.Count, limit) + 1):
        try:
            mail = items.Item(index)
        except Exception:
            continue

        try:
            key = str(mail.EntryID)
        except Exception:
            continue

        if seen.has(key):
            continue

        subject = ""
        try:
            subject = str(mail.Subject)
        except Exception:
            pass

        text = format_message(mail, cfg)

        if test_mode:
            log("ТЕСТОВЫЙ РЕЖИМ - не отправлено. Было бы отправлено:\n%s\n---" % text)
            seen.add(key)
            handled += 1
            continue

        try:
            msg_id = client.send_text(text)
            log("ОТПРАВЛЕНО (id=%s): %s" % (msg_id, subject))

            if send_files:
                for path in save_attachments(mail, tmp_dir, log):
                    try:
                        fid = client.send_file(path)
                        log("  файл отправлен (id=%s): %s" % (fid, os.path.basename(path)))
                    except TdmError as e:
                        log("  файл НЕ отправлен (%s): %s" % (os.path.basename(path), e))
                    finally:
                        try:
                            os.remove(path)
                        except OSError:
                            pass

            seen.add(key)
            handled += 1
        except TdmError as e:
            log("ОШИБКА ОТПРАВКИ (%s): %s" % (subject, e))
            # не помечаем как обработанное - попробуем в следующий раз
            break

    log("Готово. Обработано новых писем: %d" % handled)
    return handled


# =====================================================================
#  Самопроверка: поднимаем поддельный сервер и проверяем всю цепочку
# =====================================================================

def selftest():
    """Проверяет программу целиком без токена, без интернета и без Outlook."""
    import threading
    import tempfile
    from http.server import BaseHTTPRequestHandler, HTTPServer

    print("=" * 60)
    print(" САМОПРОВЕРКА mail2tdm")
    print(" Не требует ни токена, ни интернета, ни Outlook.")
    print("=" * 60)

    received = {}
    TOKEN = "SELFTEST-TOKEN"

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def _reply(self, code, obj):
            body = json.dumps(obj).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_POST(self):
            if self.headers.get("Authorization") != TOKEN:
                return self._reply(401, {"error": "bad token"})
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length) if length else b""

            if self.path.endswith("/groups/getAllUserGroupStates"):
                return self._reply(200, [{"groupId": 2387045628266976},
                                         {"groupId": 1111111111111111}])

            if "/messages/sendTextMessage/" in self.path:
                data = json.loads(raw.decode("utf-8"))
                received["text"] = data.get("message")
                received["text_path"] = self.path
                return self._reply(200, {"messageId": 111})

            if self.path.endswith("/api/v1/upload/secret/encryptable"):
                received["upload_bytes"] = raw
                received["upload_ctype"] = self.headers.get("Content-Type", "")
                return self._reply(200, {"resource": {
                    "id": "test-uuid", "url": "https://example/f.txt",
                    "key": "KEY==", "transformation": "GOST3412-2015/ECB/PKCS7Padding"}})

            if "/messages/sendFile/" in self.path:
                received["file_payload"] = json.loads(raw.decode("utf-8"))
                return self._reply(200, {"messageId": 222})

            return self._reply(404, {"error": "not found"})

    server = HTTPServer(("127.0.0.1", 0), Handler)
    port = server.server_address[1]
    threading.Thread(target=server.serve_forever, daemon=True).start()

    base = "http://127.0.0.1:%d" % port
    client = TdmClient(base, TOKEN, workspace_id="555", group_id="777",
                       log=print, min_interval=0)

    failures = []

    def check(name, condition, detail=""):
        if condition:
            print("  [OK]    %s" % name)
        else:
            print("  [ОШИБКА] %s %s" % (name, detail))
            failures.append(name)

    print("\n1. Список групп бота")
    groups = client.get_groups()
    check("сервер вернул список групп", len(groups) == 2, str(groups))
    check("ID группы прочитан", groups[0].get("groupId") == 2387045628266976)

    print("\n2. Отправка текста с русскими буквами")
    russian = 'Задача №5: проверка «кавычек», перевод\nстроки и слэш \\ '
    mid = client.send_text(russian)
    check("сервер принял сообщение", mid == 111)
    check("русский текст дошёл без искажений", received.get("text") == russian,
          repr(received.get("text")))
    check("адрес запроса верный",
          received.get("text_path") == "/botapi/v1/messages/sendTextMessage/555/777",
          received.get("text_path"))

    print("\n3. Отправка файла")
    tmp = tempfile.NamedTemporaryFile(suffix="_смета.txt", delete=False, mode="wb")
    tmp.write("содержимое файла".encode("utf-8"))
    tmp.close()
    fid = client.send_file(tmp.name, message="файл из письма")
    check("сервер принял файл", fid == 222)
    check("файл долетел целиком",
          "содержимое файла".encode("utf-8") in received.get("upload_bytes", b""))
    check("формат загрузки multipart",
          "multipart/form-data" in received.get("upload_ctype", ""))
    payload = received.get("file_payload", {})
    check("в сообщении есть ссылка на файл",
          payload.get("file", {}).get("resourceRef", {}).get("id") == "test-uuid")
    check("имя файла передано",
          payload.get("file", {}).get("fileName", "").endswith("_смета.txt"),
          payload.get("file", {}).get("fileName"))
    os.remove(tmp.name)

    print("\n4. Поведение при неверном токене")
    bad = TdmClient(base, "WRONG", workspace_id="1", group_id="2", min_interval=0)
    try:
        bad.send_text("нельзя")
        check("неверный токен отклонён", False, "ошибки не было")
    except TdmError as e:
        check("неверный токен отклонён с понятным текстом", "токен" in str(e).lower(), str(e))

    server.shutdown()

    print("\n" + "=" * 60)
    if failures:
        print(" ИТОГ: ЕСТЬ ПРОБЛЕМЫ (%d). Пришли этот вывод разработчику." % len(failures))
        return 1
    print(" ИТОГ: ВСЁ ИСПРАВНО. Программа готова к настройке.")
    print("=" * 60)
    return 0


# =====================================================================
#  Точка входа
# =====================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Пересылка писем из Outlook в канал TDM")
    parser.add_argument("--config", default=None, help="путь к файлу настроек")
    parser.add_argument("--selftest", action="store_true",
                        help="самопроверка без токена и без сети")
    parser.add_argument("--check", action="store_true",
                        help="проверить связь и показать каналы бота")
    parser.add_argument("--test-message", action="store_true",
                        help="отправить одно тестовое сообщение")
    parser.add_argument("--once", action="store_true",
                        help="разобрать новые письма один раз")
    parser.add_argument("--watch", action="store_true",
                        help="работать постоянно")
    args = parser.parse_args()

    if args.selftest:
        return selftest()

    config_path = find_config(args.config)
    cfg = read_config(config_path)

    log_path = cfg.get("LogPath", "").strip()
    if not log_path:
        log_path = os.path.join(os.path.dirname(os.path.abspath(config_path)),
                                "mail2tdm.log")
    log = Logger(log_path)

    if not cfg:
        log("НЕ НАЙДЕН файл настроек: %s" % config_path)
        log("Скопируй config.example.ini в config.ini и заполни его.")
        return 2

    test_mode = cfg.get("TestMode", "1").strip() != "0"

    client = TdmClient(
        base_url=cfg.get("BaseUrl", ""),
        token=cfg.get("Token", ""),
        workspace_id=cfg.get("WorkspaceId", ""),
        group_id=cfg.get("GroupId", ""),
        log=log,
    )

    log("Настройки: %s" % config_path)
    log("Журнал: %s" % log_path)
    log("Режим: %s" % ("ТЕСТОВЫЙ (ничего не отправляется)" if test_mode
                       else "БОЕВОЙ (сообщения уходят в канал)"))

    try:
        if args.check:
            groups = client.get_groups()
            log("Связь есть. Бот состоит в группах: %d" % len(groups))
            for g in groups[:50]:
                if isinstance(g, dict) and g.get("groupId"):
                    log("   groupId = %s" % g["groupId"])
            if not groups:
                log("Бот пока никуда не добавлен. Добавь его в нужный канал.")
            return 0

        if args.test_message:
            if test_mode:
                log("Сейчас ТЕСТОВЫЙ режим - реальная отправка отключена.")
                log("Чтобы отправить по-настоящему, поставь TestMode=0 в настройках.")
                return 0
            mid = client.send_text("Проверка связи: mail2tdm подключён.")
            log("Отправлено. Номер сообщения: %s" % mid)
            return 0

        seen_path = os.path.join(os.path.dirname(log_path), "mail2tdm_seen.json")
        seen = SeenStore(seen_path, log)

        if args.once:
            process_once(cfg, client, log, seen, test_mode)
            return 0

        if args.watch:
            interval = int(cfg.get("IntervalSeconds", "60") or 60)
            log("Работаю постоянно, проверка каждые %d сек. Ctrl+C - остановить." % interval)
            while True:
                try:
                    process_once(cfg, client, log, seen, test_mode)
                except TdmError as e:
                    log("Ошибка: %s" % e)
                except Exception as e:
                    log("Неожиданный сбой: %s" % e)
                    log(traceback.format_exc())
                time.sleep(interval)

        parser.print_help()
        return 0

    except TdmError as e:
        log("ОШИБКА: %s" % e)
        return 1
    except KeyboardInterrupt:
        log("Остановлено вручную.")
        return 0
    except Exception as e:
        log("НЕОЖИДАННЫЙ СБОЙ: %s" % e)
        log(traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
