import argparse
import logging
import os
import sqlite3
import time
from datetime import datetime
from pathlib import Path

from outlook_com import OutlookMailbox
from outlook_config import Route, Settings, load_settings
from outlook_storage import InsufficientDiskSpace, Storage, worker_lock


LOGGER = logging.getLogger("outlook_exporter")
EXPORTED = "Выгружено"
DUPLICATE = "Дубликат"
NO_ATTACHMENTS = "Нет файлов"
NO_ALLOWED_FILES = "Нет нужных файлов"
EXPORT_ERROR = "Ошибка выгрузки"
UNSAFE_FILENAME_CHARACTERS = str.maketrans({character: "_" for character in '\\/:*?"<>|'})


def attachment_filename(received: datetime, original: str) -> str:
    name = original.translate(UNSAFE_FILENAME_CHARACTERS)
    name = "".join("_" if ord(character) < 32 else character for character in name)
    name = name.strip().rstrip(". ") or "file"
    name = name.encode("utf-16-le")[:280].decode("utf-16-le", errors="ignore").rstrip(". ")
    return f"{received:%Y-%m-%d_%H-%M-%S}_{name}"


def planned_export(mailbox: OutlookMailbox, message: dict, route: Route) -> tuple[str, list[dict]]:
    attachments = list(mailbox.iter_attachments(message["id"]))
    if not attachments:
        return NO_ATTACHMENTS, []

    allowed = []
    for attachment in attachments:
        if not attachment["is_file"]:
            continue
        _, _, extension = attachment["name"].rpartition(".")
        if not route.allowed_extensions or extension.strip().casefold() in route.allowed_extensions:
            allowed.append(attachment)
    return (EXPORTED if allowed else NO_ALLOWED_FILES), allowed


def export_message(
    mailbox: OutlookMailbox,
    storage: Storage | None,
    route: Route,
    route_key: str,
    storage_id: str,
    message: dict,
    *,
    move_planned: bool,
) -> bool:
    """Save the requested files and record the result in Outlook."""
    message_id = message["id"]
    status, attachments = planned_export(mailbox, message, route) if route.save_attachments else (NO_ATTACHMENTS, [])
    message_filename = None
    if route.save_message:
        message_filename = attachment_filename(message["receivedDateTime"], "Письмо.msg")
        status = EXPORTED
    file_count = len(attachments) + int(message_filename is not None)

    if storage is None:
        LOGGER.info("Пробный запуск: %s, файлов %d%s", status, file_count,
                    ", письмо будет перемещено" if file_count and move_planned else "")
        for attachment in attachments:
            LOGGER.info("  %s", route.output_dir / attachment_filename(message["receivedDateTime"], attachment["name"]))
        if message_filename:
            LOGGER.info("  %s", route.output_dir / message_filename)
        return status == EXPORTED

    for attachment in attachments:
        attachment_id = attachment["id"]
        path = storage.save_attachment(
            route_key, storage_id, attachment_id, route.output_dir,
            attachment_filename(message["receivedDateTime"], attachment["name"]),
            lambda output: mailbox.save_attachment(message_id, attachment_id, output),
            estimated_size=attachment["size"],
        )
        LOGGER.info("Файл выгружен: %s", path)

    if message_filename:
        path = storage.save_attachment(
            route_key, storage_id, "message", route.output_dir, message_filename,
            lambda output: mailbox.save_message(message_id, output),
            temp_suffix=".msg", estimated_size=mailbox.message_size(message_id),
        )
        LOGGER.info("Письмо сохранено: %s", path)

    if file_count and not storage.exported_files_present(route_key, storage_id):
        raise OSError("Не все файлы из журнала выгрузки доступны на диске")
    mailbox.set_result(
        message_id, status, processed=True,
        mark_read=route.mark_read_after_export and status == EXPORTED,
    )
    LOGGER.info("%s, файлов %d", status, file_count)
    return status == EXPORTED


def is_exported(mailbox: OutlookMailbox, message: dict) -> bool:
    return mailbox.is_processed(message) and any(
        category.casefold() == EXPORTED.casefold() for category in message["categories"]
    )


def prepare_message(
    mailbox: OutlookMailbox,
    storage: Storage | None,
    route: Route,
    route_key: str,
    destination: str | None,
    message: dict,
    identity_cache: dict[str, str | None],
    previewed: set[str],
    *,
    trust_destination_status: bool,
) -> tuple[str | None, str] | None:
    """Export or recognize a duplicate. Return the identity and journal key for a pending move."""
    message_id = message["id"]
    pending = mailbox.is_processed(message)
    if pending:
        if not destination or storage is None or not is_exported(mailbox, message):
            return None
    else:
        if route.unread_only and not message["unread"]:
            return None
        if route.subject_contains and route.subject_contains.casefold() not in message["subject"].casefold():
            return None

    identity = None
    if destination:
        identity = identity_cache[message_id] if message_id in identity_cache else mailbox.message_identity(message_id)
    if storage is None and identity in previewed:
        LOGGER.info("Пробный запуск: копия письма, уже запланированного к выгрузке")
        return None

    storage_id = storage.export_key(route_key, message_id, identity) if storage is not None else message_id
    has_history = storage is not None and storage.has_attachment_records(route_key, storage_id)
    if pending and not has_history:
        return None

    matches = mailbox.find_destination_matches(destination, identity) if destination and identity else []
    files_missing = has_history and not storage.exported_files_present(route_key, storage_id)
    if trust_destination_status and any(is_exported(mailbox, match) for match in matches) and not files_missing:
        if storage is not None:
            mailbox.set_result(message_id, DUPLICATE, processed=True)
        LOGGER.info("Дубликат: письмо уже выгружено и есть в папке обработки")
        return None

    if not pending or files_missing:
        saved = export_message(
            mailbox, storage, route, route_key, storage_id, message,
            move_planned=bool(destination and not matches),
        )
        if not saved:
            return None

    if storage is None:
        if identity:
            previewed.add(identity)
        return None

    if destination:
        return identity, storage_id
    return None


def finish_message(
    mailbox: OutlookMailbox,
    storage: Storage,
    route_key: str,
    message_id: str,
    source: str,
    destination: str,
    identity: str | None,
    storage_id: str,
    *,
    mark_read: bool,
) -> bool:
    """Finish a successful export. Return True if the current message was moved."""
    if not storage.exported_files_present(route_key, storage_id):
        raise OSError("Файлы выгрузки отсутствуют на диске. Письмо оставлено в исходной папке")

    # Check again because Outlook could receive or move a copy during the export.
    matches = mailbox.find_destination_matches(destination, identity) if identity else []
    if not matches:
        mailbox.move_message(message_id, destination, expected_source=source)
        return True

    if any(match["id"] == message_id for match in matches):
        return False
    if not any(is_exported(mailbox, match) for match in matches):
        mailbox.set_result(matches[0]["id"], EXPORTED, processed=True, mark_read=mark_read)
    mailbox.set_result(message_id, DUPLICATE, processed=True)
    LOGGER.info("Файлы сохранены. Копия уже есть в папке обработки, исходное письмо отмечено как дубликат")
    return False


def resolve_routes(mailbox: OutlookMailbox, settings: Settings) -> list[tuple[Route, str, str | None]]:
    resolved_routes = []
    sources = set()
    for route in settings.routes:
        source = mailbox.resolve_folder(route.source)
        if source in sources:
            raise ValueError(f"Папка {' / '.join(route.source)} указана в нескольких правилах")
        sources.add(source)
        destination = mailbox.resolve_destination_folder(route.move_to) if route.move_to else None
        if destination == source:
            raise ValueError(f"Источник и назначение совпадают: {' / '.join(route.source)}")
        resolved_routes.append((route, source, destination))
    return resolved_routes


def scan(mailbox: OutlookMailbox, settings: Settings, storage: Storage | None) -> int:
    resolved_routes = resolve_routes(mailbox, settings)
    if storage is not None and settings.manage_categories:
        mailbox.ensure_categories()

    errors = 0
    outputs_by_destination = {}
    for route, _, destination in resolved_routes:
        if destination:
            output = (route.output_dir, route.save_attachments, route.save_message, route.allowed_extensions)
            outputs_by_destination.setdefault(destination, set()).add(output)
    for route, source, destination in resolved_routes:
        folder_name = " / ".join(route.source)
        messages = list(mailbox.iter_messages(source))
        LOGGER.info("%s: писем для проверки %d", folder_name, len(messages))
        route_key = f"{source}\n{route.output_dir}"
        identities = {}
        unavailable = set()
        if storage is not None and destination:
            # Reuse journals from before identity-based deduplication, even if a copy is enumerated first.
            for message in messages:
                message_id = message["id"]
                if storage.has_attachment_records(route_key, message_id):
                    try:
                        identity = mailbox.message_identity(message_id)
                        identities[message_id] = identity
                        storage.export_key(route_key, message_id, identity)
                    except (OSError, RuntimeError, ValueError, sqlite3.Error) as exc:
                        unavailable.add(message_id)
                        errors += 1
                        LOGGER.error("%s: не удалось сверить предыдущую выгрузку: %s", folder_name, exc)
        previewed = set()
        trust_destination_status = len(outputs_by_destination.get(destination, ())) <= 1
        for message in messages:
            if message["id"] in unavailable:
                continue
            try:
                prepared = prepare_message(
                    mailbox, storage, route, route_key, destination, message, identities, previewed,
                    trust_destination_status=trust_destination_status,
                )
            except InsufficientDiskSpace:
                raise
            except (OSError, RuntimeError, ValueError, sqlite3.Error) as exc:
                errors += 1
                LOGGER.error("%s: %s", folder_name, exc)
                if storage is not None and not mailbox.is_processed(message):
                    try:
                        mailbox.set_result(message["id"], EXPORT_ERROR, processed=False)
                    except (OSError, RuntimeError, ValueError) as category_error:
                        LOGGER.error("Не удалось отметить ошибку в Outlook: %s", category_error)
                continue

            if destination and storage is not None and prepared is not None:
                identity, storage_id = prepared
                try:
                    if finish_message(
                        mailbox, storage, route_key, message["id"], source, destination, identity, storage_id,
                        mark_read=route.mark_read_after_export,
                    ):
                        LOGGER.info("Письмо перемещено в %s", " / ".join(route.move_to))
                except (OSError, RuntimeError, ValueError, sqlite3.Error) as exc:
                    errors += 1
                    LOGGER.error("Не удалось завершить обработку в Outlook: %s", exc)
    return errors


def run(settings: Settings, *, once: bool, dry_run: bool) -> int:
    if dry_run:
        return run_scans(settings, None, once=True)
    with worker_lock(settings.state_db):
        storage = Storage(settings.state_db)
        try:
            return run_scans(settings, storage, once=once)
        finally:
            storage.close()


def run_scans(settings: Settings, storage: Storage | None, *, once: bool) -> int:
    while True:
        try:
            mailbox = OutlookMailbox(dry_run=storage is None, manage_categories=settings.manage_categories)
            errors = scan(mailbox, settings, storage)
        except InsufficientDiskSpace as exc:
            LOGGER.warning("Выгрузка приостановлена: %s", exc)
            errors = 1
        except (OSError, RuntimeError, sqlite3.Error) as exc:
            LOGGER.error("Не удалось проверить Outlook: %s", exc)
            errors = 1
        if once:
            return int(errors > 0)
        time.sleep(settings.poll_interval_seconds)


def main() -> int:
    parser = argparse.ArgumentParser(description="Фоновая выгрузка вложений из папок Outlook")
    parser.add_argument("--config", type=Path, default=Path("config.toml"), help="Путь к конфигурации TOML")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--once", action="store_true", help="Обработать папки один раз")
    mode.add_argument("--dry-run", action="store_true", help="Показать план, не меняя письма и файлы")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if os.name != "nt":
        LOGGER.error("Для подключения нужен классический Outlook в Windows")
        return 2
    try:
        settings = load_settings(args.config)
        return run(settings, once=args.once, dry_run=args.dry_run)
    except KeyboardInterrupt:
        LOGGER.info("Выгрузка остановлена")
        return 0
    except (OSError, RuntimeError, ValueError, sqlite3.Error) as exc:
        LOGGER.error("%s", exc)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
