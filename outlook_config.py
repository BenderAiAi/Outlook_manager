import json
import os
import tempfile
from dataclasses import dataclass
from email.errors import HeaderParseError
from email.headerregistry import Address
from pathlib import Path

try:
    import tomllib
except ImportError:
    import tomli as tomllib


DEFAULT_EXTENSIONS = ["xlsx", "xls", "xlsm", "csv", "pdf", "zip", "txt", "doc", "docx"]
DEFAULT_STATE_DB = ".outlook-exporter.sqlite3"
DEFAULT_POLL_INTERVAL = 30
DEFAULT_INBOX = ("Inbox",)
ROUTE_PATHS = ("source", "move_to", "pending", "processed", "error")
ROUTE_FLAGS = (
    "save_attachments", "save_message", "unread_only", "mark_read_after_export", "include_existing",
)
ROUTE_FILTERS = ("subject_contains", "subject_exact", "sender", "attachment_filter")
ROUTE_FIELDS = (*ROUTE_PATHS, *ROUTE_FLAGS, *ROUTE_FILTERS, "output_dir", "allowed_extensions")


@dataclass(frozen=True)
class Route:
    source: tuple[str, ...]
    output_dir: Path
    allowed_extensions: frozenset[str]
    move_to: tuple[str, ...] | None
    save_attachments: bool = True
    save_message: bool = False
    subject_contains: str = ""
    unread_only: bool = False
    mark_read_after_export: bool = False
    pending: tuple[str, ...] | None = None
    processed: tuple[str, ...] | None = None
    error: tuple[str, ...] | None = None
    subject_exact: str = ""
    sender: str = ""
    attachment_filter: str = "any"
    include_existing: bool = False


@dataclass(frozen=True)
class Settings:
    poll_interval_seconds: int
    manage_categories: bool
    state_db: Path
    routes: tuple[Route, ...]


def _unquote_path(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] == '"':
        return value[1:-1]
    return value


def folder_path(value: object, setting: str, *, allow_root: bool = False) -> tuple[str, ...]:
    if isinstance(value, str):
        path = _unquote_path(value)
        if path.startswith("\\\\"):
            value = ["", *path[2:].rstrip("\\").split("\\")]
        else:
            value = [part.strip() for part in path.replace("\\", "/").split("/")]
    if not isinstance(value, (list, tuple)) or (not value and not allow_root):
        raise ValueError(f"{setting}: укажите путь к папке Outlook")
    absolute = bool(value) and value[0] == ""
    names = value[1:] if absolute else value
    if (absolute and len(names) < (1 if allow_root else 2)) or any(
        not isinstance(part, str) or not part.strip() or any(ord(char) < 32 for char in part)
        for part in names
    ):
        raise ValueError(f"{setting}: укажите путь к папке Outlook")
    return tuple(value)


def _folder_key(path: tuple[str, ...]) -> tuple[str, ...]:
    parts = tuple(part.casefold() for part in path)
    if parts[0] == "входящие":
        return ("inbox", *parts[1:])
    return parts


def require_local_path(path: Path) -> None:
    if os.name != "nt":
        return

    drive = path.drive or Path.cwd().drive
    if drive.startswith("\\\\"):
        raise ValueError("Укажите путь на локальном диске Windows")

    from ctypes import windll

    drive_remote = 4
    if drive and windll.kernel32.GetDriveTypeW(drive + "\\") == drive_remote:
        raise ValueError("Укажите путь на локальном диске Windows")


def _resolve_local_path(path: Path) -> Path:
    require_local_path(path)
    resolved = path.resolve()
    require_local_path(resolved)
    return resolved


def filesystem_path(value: object, base: Path, setting: str) -> Path:
    if not isinstance(value, str):
        raise ValueError(f"{setting} должен содержать путь")
    value = _unquote_path(value)
    if not value.strip() or "\0" in value:
        raise ValueError(f"{setting} должен содержать путь")
    path = Path(os.path.expandvars(value)).expanduser()
    return _resolve_local_path(base / path)


def read_settings_data(config_path: Path) -> dict:
    with _resolve_local_path(config_path).open("rb") as config_file:
        return tomllib.load(config_file)


def validate_settings(data: dict, config_path: Path) -> Settings:
    if not isinstance(data, dict):
        raise ValueError("Настройки должны быть объектом")
    config_path = _resolve_local_path(config_path)
    interval = data.get("poll_interval_seconds", DEFAULT_POLL_INTERVAL)
    if type(interval) is not int or interval < 1:
        raise ValueError("poll_interval_seconds должен быть целым числом больше нуля")

    manage_categories = data.get("manage_categories", True)
    if type(manage_categories) is not bool:
        raise ValueError("manage_categories должен быть true или false")

    state_db = filesystem_path(
        data.get("state_db", DEFAULT_STATE_DB), config_path.parent, "state_db"
    )
    if state_db.is_relative_to(config_path) or state_db.is_dir():
        raise ValueError("state_db должен указывать на отдельный файл журнала")

    route_data = data.get("routes")
    if not isinstance(route_data, list) or not route_data:
        raise ValueError("Добавьте хотя бы одно правило выгрузки")

    routes = []
    sources: dict[tuple[str, ...], bool] = {}
    status_folders: dict[tuple[str, ...], str] = {}
    for index, values in enumerate(route_data, start=1):
        label = f"routes[{index}]"
        if not isinstance(values, dict):
            raise ValueError(f"{label} должен быть секцией TOML")
        pipeline = any(key in values for key in ("pending", "processed", "error"))
        extensions = values.get("allowed_extensions", [] if pipeline else DEFAULT_EXTENSIONS)
        if not isinstance(extensions, list) or any(
            not isinstance(extension, str) or not extension.strip().lstrip(".")
            for extension in extensions
        ):
            raise ValueError(f"{label}.allowed_extensions должен содержать список расширений")

        source = folder_path(values.get("source", DEFAULT_INBOX if pipeline else None), f"{label}.source")
        source_key = _folder_key(source)
        if source_key in sources and not (pipeline and sources[source_key]):
            raise ValueError(f"Папка {' / '.join(source)} указана в нескольких правилах")
        sources[source_key] = pipeline

        stages = {}
        if pipeline:
            for role in ("pending", "processed", "error"):
                path = folder_path(values.get(role), f"{label}.{role}")
                key = _folder_key(path)
                if key in status_folders and (role == "pending" or status_folders[key] != role):
                    raise ValueError(f"Папка {' / '.join(path)} уже указана для другого этапа или очереди")
                status_folders[key] = role
                stages[role] = path

        move_to = values.get("move_to")
        if move_to is not None:
            if pipeline:
                raise ValueError(f"{label}: для нового правила укажите папку «Обработанные» в processed")
            move_to = folder_path(move_to, f"{label}.move_to")
            if _folder_key(move_to) == source_key:
                raise ValueError(f"Источник и назначение совпадают: {' / '.join(source)}")

        save_attachments = values.get("save_attachments", True)
        save_message = values.get("save_message", False)
        unread_only = values.get("unread_only", False)
        mark_read_after_export = values.get("mark_read_after_export", False)
        include_existing = values.get("include_existing", False)
        for key, value in (
            ("save_attachments", save_attachments), ("save_message", save_message), ("unread_only", unread_only),
            ("mark_read_after_export", mark_read_after_export),
            ("include_existing", include_existing),
        ):
            if type(value) is not bool:
                raise ValueError(f"{label}.{key} должен быть true или false")
        if not (save_attachments or save_message):
            raise ValueError(f"{label}: выберите сохранение вложений, письма или обоих")

        filters = {}
        for key in ("subject_contains", "subject_exact", "sender"):
            value = values.get(key, "")
            if not isinstance(value, str) or any(ord(char) < 32 for char in value):
                raise ValueError(f"{label}.{key} должен содержать одну строку")
            filters[key] = value.strip() if pipeline else value
        attachment_filter = values.get("attachment_filter", "any")
        if attachment_filter not in ("any", "present", "absent"):
            raise ValueError(f"{label}.attachment_filter должен быть any, present или absent")
        if pipeline:
            if filters["subject_exact"] and filters["subject_contains"]:
                raise ValueError(f"{label}: выберите один способ сравнения темы")
            if not any(filters.values()) and attachment_filter == "any":
                raise ValueError(f"{label}: задайте тему, отправителя или условие по вложениям")
            if filters["sender"]:
                try:
                    address = Address(addr_spec=filters["sender"])
                except (ValueError, HeaderParseError, IndexError) as error:
                    raise ValueError(f"{label}: укажите полный адрес отправителя") from error
                if not address.username or not address.domain or address.addr_spec != filters["sender"]:
                    raise ValueError(f"{label}: укажите полный адрес отправителя")
        output_dir = filesystem_path(values.get("output_dir"), config_path.parent, f"{label}.output_dir")
        if output_dir.is_relative_to(config_path) or output_dir.is_relative_to(state_db) or output_dir.is_file():
            raise ValueError(f"{label}.output_dir должен указывать на каталог выгрузки")

        routes.append(Route(
            source=source,
            output_dir=output_dir,
            allowed_extensions=frozenset(extension.strip().lstrip(".").casefold() for extension in extensions),
            move_to=move_to,
            save_attachments=save_attachments,
            save_message=save_message,
            unread_only=unread_only,
            mark_read_after_export=mark_read_after_export,
            include_existing=include_existing,
            attachment_filter=attachment_filter,
            **stages,
            **filters,
        ))

    if conflict := sources.keys() & status_folders.keys():
        path = " / ".join(sorted(conflict)[0])
        raise ValueError(f"Папка {path} указана и как источник, и как папка обработки")
    return Settings(interval, manage_categories, state_db, tuple(routes))


def load_settings(config_path: Path) -> Settings:
    return validate_settings(read_settings_data(config_path), config_path)


def _quote(value: str) -> str:
    # JSON escapes the same characters as a TOML basic string except for DEL.
    return json.dumps(value, ensure_ascii=False).replace("\x7f", "\\u007f")


def _quote_list(values: list[str]) -> str:
    return "[" + ", ".join(_quote(value) for value in values) + "]"


def write_settings(data: dict, config_path: Path) -> None:
    config_path = _resolve_local_path(config_path)
    settings = validate_settings(data, config_path)

    lines = [
        f"poll_interval_seconds = {settings.poll_interval_seconds}",
        f"manage_categories = {str(settings.manage_categories).lower()}",
    ]
    if "state_db" in data:
        lines.append(f"state_db = {_quote(data['state_db'])}")
    for route, validated in zip(data["routes"], settings.routes):
        lines.extend((
            "",
            "[[routes]]",
            f"source = {_quote_list(validated.source)}",
            f"output_dir = {_quote(route['output_dir'])}",
            f"allowed_extensions = {_quote_list(route.get('allowed_extensions', [] if validated.pending else DEFAULT_EXTENSIONS))}",
        ))
        for key in ROUTE_PATHS[1:]:
            if path := getattr(validated, key):
                lines.append(f"{key} = {_quote_list(path)}")
        for key in ROUTE_FLAGS:
            if key in route:
                lines.append(f"{key} = {str(route[key]).lower()}")
        for key in ROUTE_FILTERS:
            if key in route:
                lines.append(f"{key} = {_quote(getattr(validated, key))}")
    contents = ("\n".join(lines) + "\n").encode("utf-8")

    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", dir=config_path.parent, prefix=".outlook-settings-", suffix=".tmp", delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            temporary_file.write(contents)
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
        os.replace(temporary_path, config_path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
