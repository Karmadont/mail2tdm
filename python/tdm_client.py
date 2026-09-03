"""
tdm_client.py - разговор с TDM Bot API.

Что умеет:
  * проверить связь и получить список групп, где состоит бот;
  * отправить текстовое сообщение в группу/канал;
  * загрузить файл в хранилище и отправить его сообщением.

ВАЖНО: здесь используются только средства, входящие в сам Python.
Ничего доустанавливать из интернета не нужно.

Токен нигде в этом файле не хранится - он приходит снаружи, из файла настроек.
"""

import json
import mimetypes
import os
import ssl
import time
import urllib.error
import urllib.request
import uuid


class TdmError(Exception):
    """Ошибка при разговоре с сервером TDM. Текст пригоден для показа человеку."""


class TdmClient:
    def __init__(self, base_url, token, workspace_id=None, group_id=None,
                 timeout=30, log=None, min_interval=1.1):
        self.base_url = (base_url or "").rstrip("/")
        self.token = token or ""
        self.workspace_id = str(workspace_id or "")
        self.group_id = str(group_id or "")
        self.timeout = timeout
        self.log = log or (lambda msg: None)
        # Сервер разрешает 60 запросов в минуту - держим паузу между запросами
        self.min_interval = min_interval
        self._last_call = 0.0

    # ---------- внутреннее ----------

    def _wait_rate_limit(self):
        delta = time.time() - self._last_call
        if delta < self.min_interval:
            time.sleep(self.min_interval - delta)
        self._last_call = time.time()

    def _request(self, path, data=None, content_type=None, extra_headers=None):
        """Отправить POST-запрос и вернуть разобранный JSON (или None)."""
        if not self.base_url:
            raise TdmError("Не задан адрес сервера (BaseUrl) в настройках.")
        if not self.token:
            raise TdmError("Не задан токен бота (Token) в настройках.")

        self._wait_rate_limit()

        url = self.base_url + path
        headers = {"Authorization": self.token}
        if content_type:
            headers["Content-Type"] = content_type
        if extra_headers:
            headers.update(extra_headers)

        req = urllib.request.Request(url, data=data, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read()
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8", "replace")[:500]
            except Exception:
                pass
            if e.code in (401, 403):
                raise TdmError(
                    "Сервер не принял токен (код %s). Проверь, что токен скопирован "
                    "целиком и не устарел." % e.code)
            raise TdmError("Сервер вернул ошибку %s. Ответ: %s" % (e.code, body))
        except urllib.error.URLError as e:
            raise TdmError(
                "Не удалось связаться с сервером: %s\n"
                "Проверь адрес BaseUrl и что с этого компьютера есть доступ к сети TDM."
                % e.reason)

        if not raw:
            return None
        try:
            return json.loads(raw.decode("utf-8"))
        except ValueError:
            return {"_raw": raw.decode("utf-8", "replace")[:500]}

    def _json_post(self, path, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        return self._request(path, data=body,
                             content_type="application/json; charset=utf-8")

    # ---------- публичные действия ----------

    def get_groups(self):
        """Список групп/каналов, в которых состоит бот."""
        result = self._request("/botapi/v1/groups/getAllUserGroupStates",
                               data=b"", content_type="application/json; charset=utf-8")
        if result is None:
            return []
        if isinstance(result, list):
            return result
        return [result]

    def send_text(self, message, workspace_id=None, group_id=None):
        """Отправить текстовое сообщение. Возвращает messageId."""
        ws = str(workspace_id or self.workspace_id)
        gr = str(group_id or self.group_id)
        if not ws or not gr:
            raise TdmError("Не заданы WorkspaceId и/или GroupId в настройках.")

        path = "/botapi/v1/messages/sendTextMessage/%s/%s" % (ws, gr)
        result = self._json_post(path, {"message": message})
        if isinstance(result, dict):
            return result.get("messageId")
        return None

    def upload_file(self, file_path):
        """
        Загрузить файл в хранилище TDM. Шифрует сам сервер, поэтому нам
        не нужно ничего шифровать самим.
        Возвращает словарь resourceRef: id, url, key, transformation.
        """
        if not os.path.isfile(file_path):
            raise TdmError("Файл не найден: %s" % file_path)

        file_name = os.path.basename(file_path)
        mime = mimetypes.guess_type(file_name)[0] or "application/octet-stream"
        with open(file_path, "rb") as f:
            content = f.read()

        boundary = "----mail2tdm" + uuid.uuid4().hex
        body = self._build_multipart(boundary, file_name, mime, content)

        result = self._request(
            "/api/v1/upload/secret/encryptable",
            data=body,
            content_type="multipart/form-data; boundary=%s" % boundary,
            extra_headers={"Content-Length": str(len(body))},
        )

        if not isinstance(result, dict):
            raise TdmError("Хранилище вернуло неожиданный ответ при загрузке файла.")
        resource = result.get("resource")
        if not isinstance(resource, dict):
            # некоторые варианты метода отдают только url
            if result.get("url"):
                resource = {"url": result["url"]}
            else:
                raise TdmError("В ответе хранилища нет сведений о файле.")
        return resource, len(content), mime, file_name

    @staticmethod
    def _build_multipart(boundary, file_name, mime, content):
        """Собрать тело запроса формата multipart/form-data вручную."""
        # Имя файла кодируем в UTF-8: в письмах бывают русские названия
        head = (
            "--%s\r\n"
            'Content-Disposition: form-data; name="file"; filename="%s"\r\n'
            "Content-Type: %s\r\n\r\n" % (boundary, file_name, mime)
        ).encode("utf-8")
        tail = ("\r\n--%s--\r\n" % boundary).encode("utf-8")
        return head + content + tail

    def send_file(self, file_path, message="", workspace_id=None, group_id=None):
        """Загрузить файл и отправить его сообщением в группу."""
        ws = str(workspace_id or self.workspace_id)
        gr = str(group_id or self.group_id)
        if not ws or not gr:
            raise TdmError("Не заданы WorkspaceId и/или GroupId в настройках.")

        resource, length, mime, file_name = self.upload_file(file_path)

        payload = {
            "clientRandomId": int(uuid.uuid4().int % 2_000_000_000),
            "file": {
                "fileName": file_name,
                "length": length,
                "mimeType": mime,
                "resourceRef": {
                    "id": resource.get("id", ""),
                    "key": resource.get("key", ""),
                    "transformation": resource.get("transformation", ""),
                    "url": resource.get("url", ""),
                },
            },
        }
        if message:
            payload["message"] = message

        path = "/botapi/v1/messages/sendFile/%s/%s" % (ws, gr)
        result = self._json_post(path, payload)
        if isinstance(result, dict):
            return result.get("messageId")
        return None
