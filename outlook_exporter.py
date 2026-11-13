import argparse
import logging
import os
import sqlite3
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import NamedTuple

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
PROCESSING_ERRORS = (OSError, RuntimeError, ValueError, sqlite3.Error)


class PipelineRoute(NamedTuple):
    route: Route
    source: str
    pending: str | None
    processed: str | None
    error: str | None

    @property
    def key(self) -> str:
        return f"pipeline\n{self.source}\n{self.pending}"


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


def save_message_files(
    mailbox: OutlookMailbox,
    storage: Storage | None,
    route: Route,
    route_key: str,
    storage_id: str,
    message: dict,
) -> tuple[str, int]:
    message_id = message["id"]
    status, attachments = planned_export(mailbox, message, route) if route.save_attachments else (NO_ATTACHMENTS, [])
    message_filename = None
    if route.save_message:
        message_filename = attachment_filename(message["receivedDateTime"], "Письмо.msg")
        status = EXPORTED
    file_count = len(attachments) + int(message_filename is not None)

    if storage is None:
        for attachment in attachments:
            LOGGER.info("План сохранения: %s", route.output_dir / attachment_filename(message["receivedDateTime"], attachment["name"]))
        if message_filename:
            LOGGER.info("План сохранения: %s", route.output_dir / message_filename)
        return status, file_count

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
    return status, file_count


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
    """Save the requested files and record the legacy result in Outlook."""
    status, file_count = save_message_files(mailbox, storage, route, route_key, storage_id, message)
    if storage is None:
        LOGGER.info("Пробный запуск: %s, файлов %d%s", status, file_count,
                    ", письмо будет перемещено" if file_count and move_planned else "")
        return status == EXPORTED

    mailbox.set_result(
        message["id"], status, processed=True,
        mark_read=route.mark_read_after_export and status == EXPORTED,
    )
    LOGGER.info("%s, файлов %d%s", status, file_count,
                ". Письмо осталось в исходной папке по старому правилу" if not file_count else "")
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
    skipped: Counter,
    *,
    trust_destination_status: bool,
) -> tuple[str | None, str] | None:
    """Export or recognize a duplicate. Return the identity and journal key for a pending move."""
    message_id = message["id"]
    pending = mailbox.is_processed(message)
    if pending:
        if not destination or storage is None or not is_exported(mailbox, message):
            skipped["processed"] += 1
            return None
    else:
        if route.unread_only and not message["unread"]:
            skipped["read"] += 1
            return None
        if route.subject_contains and route.subject_contains.casefold() not in message["subject"].casefold():
            skipped["subject"] += 1
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
        skipped["unverified"] += 1
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
            skipped["no_files"] += 1
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
        if route.pending is not None:
            continue
        source = mailbox.resolve_folder(route.source)
        if source in sources:
            raise ValueError(f"Папка {' / '.join(route.source)} указана в нескольких правилах")
        sources.add(source)
        destination = mailbox.resolve_destination_folder(route.move_to) if route.move_to else None
        if destination == source:
            raise ValueError(f"Источник и назначение совпадают: {' / '.join(route.source)}")
        resolved_routes.append((route, source, destination))
    return resolved_routes


def resolve_pipeline_routes(
    mailbox: OutlookMailbox, settings: Settings, *, create_missing: bool,
) -> list[PipelineRoute]:
    if not any(route.pending is not None for route in settings.routes):
        return []

    source_types: dict[str, bool] = {}
    resolved = []
    for route in settings.routes:
        source = mailbox.resolve_folder(route.source)
        pipeline = route.pending is not None
        if source in source_types and not (pipeline and source_types[source]):
            raise ValueError(f"Папка {' / '.join(route.source)} указана в нескольких правилах")
        source_types[source] = pipeline
        if pipeline:
            folders = [mailbox.validate_destination_folder(getattr(route, role))
                       for role in ("pending", "processed", "error")]
            resolved.append(PipelineRoute(route, source, *folders))

    def check_roles(routes: list[PipelineRoute]) -> None:
        folder_roles = {}
        for item in routes:
            for role in ("pending", "processed", "error"):
                folder_id = getattr(item, role)
                if folder_id is None:
                    continue
                path = " / ".join(getattr(item.route, role))
                if folder_id in source_types:
                    raise ValueError(f"Папка {path} указана и как источник, и как папка обработки")
                if folder_id in folder_roles and (role == "pending" or folder_roles[folder_id] != role):
                    raise ValueError(f"Папка {path} уже указана для другого этапа или очереди")
                folder_roles[folder_id] = role

    check_roles(resolved)
    if not create_missing:
        return resolved
    prepared = [PipelineRoute(item.route, item.source, *(
        getattr(item, role) or mailbox.ensure_destination_folder(getattr(item.route, role))
        for role in ("pending", "processed", "error")
    )) for item in resolved]
    check_roles(prepared)
    return prepared


def matches_incoming(mailbox: OutlookMailbox, message: dict, route: Route) -> bool:
    if route.unread_only and not message["unread"]:
        return False
    subject = message["subject"].casefold()
    if route.subject_exact and subject != route.subject_exact.casefold():
        return False
    if route.subject_contains and route.subject_contains.casefold() not in subject:
        return False
    message_id = message["id"]
    if route.sender and mailbox.message_sender(message_id).casefold() != route.sender.casefold():
        return False
    if route.attachment_filter == "any":
        return True
    has_files = any(attachment["is_file"] for attachment in mailbox.iter_attachments(message_id))
    return has_files == (route.attachment_filter == "present")


def move_to_error(
    mailbox: OutlookMailbox, storage: Storage, item: PipelineRoute, source: str, message_id: str,
    *, subject: str = "",
) -> bool:
    label = f"письмо {subject!r}" if subject else "письмо из журнала"
    try:
        storage.queue_pipeline_error(item.key, source, message_id)
        if not mailbox.is_in_folder(message_id, source):
            LOGGER.warning("Перенос в «Ошибка» пропущен: %s уже вне исходной папки", label)
            storage.clear_pipeline_error(item.key, message_id)
            return True
        storage.discard_pipeline_claim(item.key, message_id)
        mailbox.move_message(message_id, item.error, expected_source=source)
        storage.clear_pipeline_error(item.key, message_id)
        LOGGER.info("Перенос в %s: %s", " / ".join(item.route.error), label)
        return True
    except PROCESSING_ERRORS as exc:
        LOGGER.error("Не удалось перенести %s в «Ошибка»: %s. Перенос повторится", label, exc)
        return False


def process_pending_message(
    mailbox: OutlookMailbox, storage: Storage, item: PipelineRoute, message: dict,
) -> None:
    message_id = message["id"]
    route = item.route
    identity = mailbox.message_identity(message_id, allow_unavailable=True)
    if identity:
        previous = storage.pipeline_processed(item.key, identity)
        if previous is not None:
            previous_id, moved = previous
            if moved:
                raise ValueError("Дубликат: письмо с таким Message-ID уже обработано этим правилом")
            if previous_id != message_id:
                # Recover when Outlook completed a move before the journal update.
                if mailbox.find_destination_matches(item.processed, identity, read_legacy_status=False):
                    storage.set_pipeline_processed(item.key, identity, previous_id, moved=True)
                    raise ValueError("Дубликат: письмо с таким Message-ID уже находится в папке обработки")
                if mailbox.is_in_folder(previous_id, item.pending):
                    raise ValueError("Дубликат: письмо с таким Message-ID уже ожидает переноса")

    if not mailbox.is_in_folder(message_id, item.pending):
        LOGGER.info("Письмо %r уже вне очереди. Сохранение пропущено", message["subject"])
        return

    legacy_key = f"{item.pending}\n{route.output_dir}"
    previous_export = storage.previous_export_key(legacy_key, message_id, identity)
    file_key = legacy_key if previous_export else f"{item.key}\n{route.output_dir}"
    status, count = save_message_files(mailbox, storage, route, file_key, previous_export or message_id, message)
    if route.mark_read_after_export:
        try:
            mailbox.mark_read(message_id)
        except PROCESSING_ERRORS as exc:
            LOGGER.warning("Не удалось отметить письмо %r прочитанным: %s", message["subject"], exc)

    # Journal the claim before moving so a crash after Outlook Move does not export a later copy.
    if identity:
        storage.set_pipeline_processed(item.key, identity, message_id, moved=False)
    moved_id = mailbox.move_message(message_id, item.processed, expected_source=item.pending)
    if identity:
        storage.set_pipeline_processed(item.key, identity, moved_id, moved=True)
    LOGGER.info("Письмо %r → %s. %s, файлов %d", message["subject"], " / ".join(route.processed), status, count)


def scan_pipeline(mailbox: OutlookMailbox, routes: list[PipelineRoute], storage: Storage | None) -> int:
    groups: dict[str, list[PipelineRoute]] = {}
    cutoffs: dict[str, datetime | None] = {}
    started = datetime.now().replace(microsecond=0)
    errors = 0
    for item in routes:
        groups.setdefault(item.source, []).append(item)
        cutoff = storage.pipeline_cutoff(item.key, started) if storage is not None else None
        cutoffs[item.key] = None if item.route.include_existing else cutoff

    for source, group in groups.items():
        folder_name = " / ".join(group[0].route.source)
        eligible_since = [cutoffs[item.key] for item in group]
        earliest = None if None in eligible_since else min(eligible_since)
        matched = moved = 0
        retried = set()
        try:
            if storage is not None:
                for item in group:
                    for message_id in storage.pipeline_errors_in(item.key, source):
                        retried.add(message_id)
                        errors += not move_to_error(mailbox, storage, item, source, message_id)
            messages = list(mailbox.iter_messages(source, read_legacy_status=False, received_since=earliest))
        except PROCESSING_ERRORS as exc:
            LOGGER.error("Не удалось проверить входящие в %s: %s", folder_name, exc)
            errors += 1
            continue

        for message in messages:
            message_id = message["id"]
            if message_id in retried:
                continue
            for item in group:
                cutoff = cutoffs[item.key]
                try:
                    if cutoff is not None and message["receivedDateTime"] < cutoff:
                        continue
                    if not matches_incoming(mailbox, message, item.route):
                        continue
                except PROCESSING_ERRORS as exc:
                    LOGGER.error("Письмо %r осталось в %s: не удалось проверить условие правила: %s",
                                 message["subject"], folder_name, exc)
                    errors += 1
                    break

                matched += 1
                try:
                    pending_name = " / ".join(item.route.pending)
                    if storage is None:
                        LOGGER.info("Пробный запуск: письмо %r → %s%s", message["subject"], pending_name,
                                    " (рабочий запуск начнёт отбор с даты активации)" if not item.route.include_existing else "")
                        status, count = save_message_files(mailbox, None, item.route, "", message_id, message)
                        LOGGER.info("Пробный запуск: %s, файлов %d → %s", status, count, " / ".join(item.route.processed))
                    else:
                        mailbox.move_message(message_id, item.pending, expected_source=source)
                        moved += 1
                        LOGGER.info("Письмо %r → %s", message["subject"], pending_name)
                except PROCESSING_ERRORS as exc:
                    LOGGER.error("Не удалось поставить письмо %r в очередь: %s", message["subject"], exc)
                    errors += 1
                    if storage is not None:
                        move_to_error(mailbox, storage, item, source, message_id, subject=message["subject"])
                break
        LOGGER.info("%s: проверено писем %d, совпало с правилами %d, в очередь перенесено %d",
                    folder_name, len(messages), matched, moved)

    for item in routes:
        if item.pending is None:
            continue
        folder_name = " / ".join(item.route.pending)
        try:
            messages = list(mailbox.iter_messages(item.pending, read_legacy_status=False))
        except PROCESSING_ERRORS as exc:
            LOGGER.error("Не удалось прочитать очередь %s: %s", folder_name, exc)
            errors += 1
            continue
        LOGGER.info("%s: писем в очереди %d", folder_name, len(messages))
        for message in messages:
            try:
                if storage is None:
                    status, count = save_message_files(mailbox, None, item.route, "", message["id"], message)
                    LOGGER.info("Пробный запуск: %s, файлов %d → %s", status, count, " / ".join(item.route.processed))
                elif storage.pipeline_error_pending(item.key, message["id"]):
                    errors += not move_to_error(mailbox, storage, item, item.pending, message["id"], subject=message["subject"])
                else:
                    process_pending_message(mailbox, storage, item, message)
            except InsufficientDiskSpace as exc:
                LOGGER.warning("Письмо %r ждёт в %s: %s", message["subject"], folder_name, exc)
                errors += 1
            except PROCESSING_ERRORS as exc:
                LOGGER.error("Не удалось обработать письмо %r в %s: %s", message["subject"], folder_name, exc)
                errors += 1
                if storage is not None:
                    move_to_error(mailbox, storage, item, item.pending, message["id"], subject=message["subject"])
    return errors


def scan_legacy(
    mailbox: OutlookMailbox, settings: Settings, storage: Storage | None,
    resolved_routes: list[tuple[Route, str, str | None]],
) -> int:
    if resolved_routes and storage is not None and settings.manage_categories:
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
        skipped = Counter()
        trust_destination_status = len(outputs_by_destination.get(destination, ())) <= 1
        for message in messages:
            if message["id"] in unavailable:
                continue
            try:
                prepared = prepare_message(
                    mailbox, storage, route, route_key, destination, message, identities, previewed, skipped,
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
        if skipped:
            LOGGER.info("%s: прочитано %d, тема не совпала %d, уже отмечено %d, без записей выгрузки %d, без файлов %d",
                        folder_name, skipped["read"], skipped["subject"], skipped["processed"],
                        skipped["unverified"], skipped["no_files"])
    return errors


def scan(
    mailbox: OutlookMailbox, settings: Settings, storage: Storage | None, *, watched: set[str] | None = None,
) -> int:
    legacy = resolve_routes(mailbox, settings)
    pipeline = resolve_pipeline_routes(mailbox, settings, create_missing=storage is not None)
    if watched is not None:
        for item in pipeline:
            watched.add(item.source)
            if item.pending is not None:
                watched.add(item.pending)
    errors = scan_pipeline(mailbox, pipeline, storage)
    return errors + scan_legacy(mailbox, settings, storage, legacy)


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
        mailbox = None
        watched: set[str] = set()
        try:
            mailbox = OutlookMailbox(dry_run=storage is None, manage_categories=settings.manage_categories)
            errors = scan(mailbox, settings, storage, watched=watched)
        except InsufficientDiskSpace as exc:
            LOGGER.warning("Выгрузка приостановлена: %s", exc)
            errors = 1
        except (OSError, RuntimeError, sqlite3.Error) as exc:
            LOGGER.error("Не удалось проверить Outlook: %s", exc)
            errors = 1
        if once:
            return int(errors > 0)
        if mailbox is not None and watched:
            mailbox.wait_for_changes(watched, settings.poll_interval_seconds)
        else:
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
