"""Local web editor for the Outlook Exporter configuration."""

import argparse
import hmac
import json
import os
import secrets
import sys
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from outlook_com import OutlookMailbox
from outlook_config import read_settings_data, require_local_path, validate_settings, write_settings
from outlook_exporter import resolve_routes


HOST = "127.0.0.1"
APP_DIR = Path(__file__).resolve().parent
MAX_REQUEST_BYTES = 256 * 1024


class RequestError(ValueError):
    def __init__(self, status: HTTPStatus, message: str):
        super().__init__(message)
        self.status = status


def folder_parts(value: object) -> tuple[str, ...]:
    if not isinstance(value, list) or any(
        not isinstance(part, str) or not part.strip() or any(ord(char) < 32 for char in part)
        for part in value
    ):
        raise ValueError("Укажите путь к папке Outlook как список имён")
    return tuple(value)


def open_mailbox(*, dry_run: bool) -> OutlookMailbox:
    if os.name != "nt":
        raise ValueError("Подключение к Outlook доступно на Windows с классическим Outlook")
    return OutlookMailbox(dry_run=dry_run, manage_categories=False)


def pick_output_directory() -> str | None:
    if os.name != "nt":
        raise ValueError("Системный выбор каталога доступен на Windows. Введите путь вручную.")

    from pywintypes import com_error
    from win32com.client import Dispatch

    return_only_fs_dirs = 0x01
    edit_box = 0x10
    new_dialog_style = 0x40
    try:
        folder = Dispatch("Shell.Application").BrowseForFolder(
            0, "Выберите каталог для выгрузки из Outlook", return_only_fs_dirs | edit_box | new_dialog_style
        )
        if folder is None:
            return None
        path = Path(folder.Self.Path)
    except com_error as error:
        raise ValueError("Не удалось открыть системный выбор каталога Windows") from error

    require_local_path(path)
    path = path.resolve()
    require_local_path(path)
    if not path.is_dir():
        raise ValueError("Выберите каталог на локальном диске")
    return str(path)


class SetupServer(HTTPServer):
    def __init__(self, config_path: Path, port: int):
        require_local_path(config_path)
        self.config_path = config_path.resolve()
        require_local_path(self.config_path)
        self.csrf_token = secrets.token_hex(32)
        self.page = (APP_DIR / "settings.html").read_bytes()
        super().__init__((HOST, port), SettingsHandler)
        self.authority = f"{HOST}:{self.server_port}"
        self.url = f"http://{self.authority}"

    def get_request(self):
        connection, address = super().get_request()
        connection.settimeout(5)
        return connection, address


class SettingsHandler(BaseHTTPRequestHandler):
    server: SetupServer

    def _send(self, status: HTTPStatus, content: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; "
            "connect-src 'self'; img-src 'self' data:; form-action 'none'; frame-ancestors 'none'; base-uri 'none'",
        )
        self.end_headers()
        if content:
            self.wfile.write(content)

    def _json(self, status: HTTPStatus, data: dict) -> None:
        self._send(status, json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def _body(self) -> dict:
        if self.headers.get_content_type() != "application/json":
            raise RequestError(HTTPStatus.UNSUPPORTED_MEDIA_TYPE, "Ожидается запрос JSON")
        lengths = self.headers.get_all("Content-Length", [])
        if len(lengths) != 1 or not lengths[0].isascii() or not lengths[0].isdigit():
            raise RequestError(HTTPStatus.LENGTH_REQUIRED, "Не указан размер запроса")
        length = int(lengths[0])
        if length > MAX_REQUEST_BYTES:
            raise RequestError(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "Слишком большой запрос")
        data = json.loads(self.rfile.read(length).decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Ожидается объект JSON")
        return data

    def _get_settings(self) -> dict:
        config_path = self.server.config_path
        result = {"config_path": str(config_path), "csrf_token": self.server.csrf_token, "exists": True}
        try:
            data = read_settings_data(config_path)
            validate_settings(data, config_path)
        except FileNotFoundError:
            result["exists"] = False
            data = read_settings_data(APP_DIR / "config.example.toml")
        except ValueError as error:
            result["load_error"] = f"Не удалось прочитать конфигурацию: {error}. В форме открыт пример."
            data = read_settings_data(APP_DIR / "config.example.toml")
        result["settings"] = {
            key: data[key] for key in ("poll_interval_seconds", "manage_categories", "state_db", "routes") if key in data
        }
        allowed_route_fields = (
            "source", "output_dir", "allowed_extensions", "move_to", "save_attachments",
            "save_message", "subject_contains", "unread_only", "mark_read_after_export",
        )
        result["settings"]["routes"] = [
            {key: route[key] for key in allowed_route_fields if key in route} for route in data["routes"]
        ]
        return result

    def _get(self, path: str, query: str) -> None:
        if path == "/":
            self._send(HTTPStatus.OK, self.server.page, "text/html; charset=utf-8")
        elif path == "/api/settings":
            self._json(HTTPStatus.OK, self._get_settings())
        elif path == "/api/folders":
            params = parse_qs(query, max_num_fields=1)
            parts = folder_parts(json.loads(params.get("path", ["[]"])[0]))
            folders = open_mailbox(dry_run=True).list_folders(parts)
            self._json(HTTPStatus.OK, {"folders": folders})
        elif path == "/favicon.ico":
            self._send(HTTPStatus.NO_CONTENT, b"", "image/x-icon")
        else:
            raise RequestError(HTTPStatus.NOT_FOUND, "Страница не найдена")

    def _post(self, path: str) -> None:
        token = self.headers.get("X-CSRF-Token", "")
        if self.headers.get("Origin") != self.server.url or not hmac.compare_digest(
            token.encode("utf-8"), self.server.csrf_token.encode("ascii")
        ):
            raise RequestError(HTTPStatus.FORBIDDEN, "Откройте страницу настроек заново")
        data = self._body()

        if path == "/api/settings":
            write_settings(data.get("settings"), self.server.config_path)
            result = {"message": "Настройки сохранены. Запустите Exporter заново, чтобы применить их."}
        elif path == "/api/validate":
            settings = validate_settings(data.get("settings"), self.server.config_path)
            if data.get("check_outlook") is True:
                resolve_routes(open_mailbox(dry_run=True), settings)
                result = {"message": "Настройки корректны. Все папки найдены в текущем Outlook."}
            else:
                result = {"message": "Настройки корректны. Наличие папок Outlook можно проверить отдельно."}
        elif path == "/api/folders":
            parent = folder_parts(data.get("parent"))
            name = data.get("name")
            if not isinstance(name, str) or not name.strip():
                raise ValueError("Введите имя новой папки")
            result = open_mailbox(dry_run=False).create_folder(parent, name)
        elif path == "/api/outlook-rules":
            folder = folder_parts(data.get("folder"))
            if not folder:
                raise ValueError("Выберите папку, в которую Outlook будет складывать письма")
            sender = data.get("sender")
            subject = data.get("subject_contains")
            if not isinstance(sender, str) or not sender.strip():
                raise ValueError("Введите адрес отправителя")
            if not isinstance(subject, str) or not subject.strip():
                raise ValueError("Введите постоянную часть темы")
            result = open_mailbox(dry_run=False).create_incoming_rule(folder, sender.strip(), subject.strip())
            name = result["name"]
            if result["created"]:
                result["message"] = f"Правило «{name}» создано в Outlook. Оно применяется к новым письмам."
            else:
                result["message"] = f"Правило «{name}» уже есть. Его состояние можно проверить в Outlook."
        elif path == "/api/output-directory":
            result = {"path": pick_output_directory()}
        else:
            raise RequestError(HTTPStatus.NOT_FOUND, "Страница не найдена")
        self._json(HTTPStatus.OK, result)

    def _handle(self, *, post: bool) -> None:
        try:
            if self.headers.get_all("Host") != [self.server.authority]:
                raise RequestError(HTTPStatus.FORBIDDEN, "Откройте локальный адрес из терминала")
            target = urlsplit(self.path)
            if post:
                self._post(target.path)
            else:
                self._get(target.path, target.query)
        except RequestError as error:
            self._json(error.status, {"error": str(error)})
        except ImportError:
            self._json(HTTPStatus.BAD_REQUEST, {"error": "На этом компьютере нужен pywin32 из requirements.txt"})
        except (ValueError, RuntimeError, OSError) as error:
            self._json(HTTPStatus.BAD_REQUEST, {"error": str(error)})

    def do_GET(self) -> None:
        self._handle(post=False)

    def do_POST(self) -> None:
        self._handle(post=True)

    def log_message(self, *_args) -> None:
        pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Локальная настройка Outlook Exporter")
    parser.add_argument("--config", type=Path, default=Path("config.toml"), help="Путь к файлу настроек TOML")
    parser.add_argument("--port", type=int, default=8765, help="Локальный порт")
    parser.add_argument("--no-browser", action="store_true", help="Показать адрес, не открывая браузер")
    args = parser.parse_args()

    try:
        with SetupServer(args.config, args.port) as server:
            print(f"Настройки Outlook Exporter: {server.url}", flush=True)
            print("Для остановки нажмите Ctrl+C.", flush=True)
            if not args.no_browser:
                webbrowser.open(server.url)
            server.serve_forever()
    except KeyboardInterrupt:
        return 0
    except (OSError, ValueError, OverflowError) as error:
        print(f"Не удалось запустить настройки: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
