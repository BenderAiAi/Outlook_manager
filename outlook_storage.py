"""Publish attachment files without overwriting existing files."""

import errno
import hashlib
import logging
import os
import shutil
import sqlite3
import stat
import tempfile
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path


LOGGER = logging.getLogger(__name__)
DISK_RESERVE_BYTES = 256 * 1024 * 1024


class InsufficientDiskSpace(OSError):
    pass


@contextmanager
def worker_lock(database: Path) -> Iterator[None]:
    database = database.resolve()
    database.parent.mkdir(parents=True, exist_ok=True)
    lock_path = database.with_name(f"{database.name}.lock")
    with lock_path.open("a+b") as lock_file:
        try:
            if os.name == "nt":
                import msvcrt

                lock_file.seek(0)
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as error:
            if error.errno not in (errno.EACCES, errno.EAGAIN, errno.EDEADLK):
                raise
            raise RuntimeError(f"Выгрузка с журналом {database} уже запущена") from error
        # Closing the handle releases the lock while retaining the lock file.
        yield


class Storage:
    def __init__(self, database: Path):
        database.parent.mkdir(parents=True, exist_ok=True)
        self._connection = sqlite3.connect(database, timeout=30)
        self._connection.execute("""
            CREATE TABLE IF NOT EXISTS attachment_files (
                route_key TEXT NOT NULL,
                message_id TEXT NOT NULL,
                attachment_id TEXT NOT NULL,
                requested_path TEXT NOT NULL,
                collision_number INTEGER NOT NULL DEFAULT 0,
                staged_path TEXT,
                PRIMARY KEY (route_key, message_id, attachment_id)
            )
        """)
        self._connection.execute("""
            CREATE TABLE IF NOT EXISTS message_export_keys (
                route_key TEXT NOT NULL,
                identity_hash TEXT NOT NULL,
                storage_message_id TEXT NOT NULL,
                PRIMARY KEY (route_key, identity_hash)
            )
        """)
        self._connection.execute("""
            CREATE TABLE IF NOT EXISTS pipeline_cutoffs (
                route_key TEXT PRIMARY KEY,
                activated_at TEXT NOT NULL
            )
        """)
        self._connection.execute("""
            CREATE TABLE IF NOT EXISTS pipeline_processed (
                route_key TEXT NOT NULL,
                identity_hash TEXT NOT NULL,
                message_id TEXT NOT NULL,
                moved INTEGER NOT NULL,
                PRIMARY KEY (route_key, identity_hash)
            )
        """)
        self._connection.execute("""
            CREATE TABLE IF NOT EXISTS pipeline_errors (
                route_key TEXT NOT NULL,
                source_id TEXT NOT NULL,
                message_id TEXT NOT NULL,
                PRIMARY KEY (route_key, message_id)
            )
        """)

    def close(self) -> None:
        self._connection.close()

    def pipeline_cutoff(self, route_key: str, started: datetime) -> datetime:
        with self._connection:
            self._connection.execute(
                "INSERT OR IGNORE INTO pipeline_cutoffs VALUES (?, ?)", (route_key, started.isoformat()),
            )
            (value,) = self._connection.execute(
                "SELECT activated_at FROM pipeline_cutoffs WHERE route_key = ?", (route_key,),
            ).fetchone()
        return datetime.fromisoformat(value)

    def pipeline_processed(self, route_key: str, identity: str) -> tuple[str, bool] | None:
        identity_hash = hashlib.sha256(identity.encode("utf-8")).hexdigest()
        row = self._connection.execute(
            "SELECT message_id, moved FROM pipeline_processed WHERE route_key = ? AND identity_hash = ?",
            (route_key, identity_hash),
        ).fetchone()
        return (row[0], bool(row[1])) if row else None

    def set_pipeline_processed(self, route_key: str, identity: str, message_id: str, *, moved: bool) -> None:
        identity_hash = hashlib.sha256(identity.encode("utf-8")).hexdigest()
        with self._connection:
            self._connection.execute(
                """
                INSERT INTO pipeline_processed VALUES (?, ?, ?, ?)
                ON CONFLICT (route_key, identity_hash)
                DO UPDATE SET message_id = excluded.message_id, moved = excluded.moved
                """,
                (route_key, identity_hash, message_id, moved),
            )

    def pipeline_error_pending(self, route_key: str, message_id: str) -> bool:
        return self._connection.execute(
            "SELECT 1 FROM pipeline_errors WHERE route_key = ? AND message_id = ?", (route_key, message_id),
        ).fetchone() is not None

    def pipeline_errors_in(self, route_key: str, source_id: str) -> list[str]:
        return [row[0] for row in self._connection.execute(
            "SELECT message_id FROM pipeline_errors WHERE route_key = ? AND source_id = ?",
            (route_key, source_id),
        )]

    def queue_pipeline_error(self, route_key: str, source_id: str, message_id: str) -> None:
        with self._connection:
            self._connection.execute(
                "INSERT OR IGNORE INTO pipeline_errors (route_key, source_id, message_id) VALUES (?, ?, ?)",
                (route_key, source_id, message_id),
            )

    def discard_pipeline_claim(self, route_key: str, message_id: str) -> None:
        with self._connection:
            self._connection.execute(
                "DELETE FROM pipeline_processed WHERE route_key = ? AND message_id = ? AND moved = 0",
                (route_key, message_id),
            )

    def clear_pipeline_error(self, route_key: str, message_id: str) -> None:
        with self._connection:
            self._connection.execute("DELETE FROM pipeline_errors WHERE route_key = ? AND message_id = ?", (route_key, message_id))

    def has_saved_attachments(self, route_key: str, message_id: str) -> bool:
        return self._connection.execute(
            """
            SELECT 1 FROM attachment_files
            WHERE route_key = ? AND message_id = ? AND staged_path IS NULL
            LIMIT 1
            """,
            (route_key, message_id),
        ).fetchone() is not None

    def has_attachment_records(self, route_key: str, message_id: str) -> bool:
        return self._connection.execute(
            "SELECT 1 FROM attachment_files WHERE route_key = ? AND message_id = ? LIMIT 1",
            (route_key, message_id),
        ).fetchone() is not None

    def previous_export_key(self, route_key: str, message_id: str, identity: str | None) -> str | None:
        if self.has_attachment_records(route_key, message_id):
            return message_id
        if identity is None:
            return None
        identity_hash = hashlib.sha256(identity.encode("utf-8")).hexdigest()
        row = self._connection.execute(
            "SELECT storage_message_id FROM message_export_keys WHERE route_key = ? AND identity_hash = ?",
            (route_key, identity_hash),
        ).fetchone()
        return row[0] if row and self.has_attachment_records(route_key, row[0]) else None

    def export_key(self, route_key: str, message_id: str, identity: str | None) -> str:
        if identity is None:
            return message_id

        identity_hash = hashlib.sha256(identity.encode("utf-8")).hexdigest()
        with self._connection:
            self._connection.execute("BEGIN IMMEDIATE")
            existing = self._connection.execute(
                "SELECT storage_message_id FROM message_export_keys WHERE route_key = ? AND identity_hash = ?",
                (route_key, identity_hash),
            ).fetchone()
            if existing is not None:
                return existing[0]

            storage_id = message_id if self.has_attachment_records(route_key, message_id) else f"internet:{identity_hash}"
            self._connection.execute(
                "INSERT INTO message_export_keys (route_key, identity_hash, storage_message_id) VALUES (?, ?, ?)",
                (route_key, identity_hash, storage_id),
            )
            return storage_id

    def exported_files_present(self, route_key: str, message_id: str) -> bool:
        rows = self._connection.execute(
            """
            SELECT requested_path, collision_number, staged_path FROM attachment_files
            WHERE route_key = ? AND message_id = ?
            """,
            (route_key, message_id),
        ).fetchall()
        return bool(rows) and all(
            staged_path is None and self._is_regular_file(self._destination(Path(original), collision_number))
            for original, collision_number, staged_path in rows
        )

    def save_attachment(
        self,
        route_key: str,
        message_id: str,
        attachment_id: str,
        directory: Path,
        filename: str,
        download: Callable[[Path], None],
        *,
        temp_suffix: str = ".tmp",
        estimated_size: int = 0,
    ) -> Path:
        if filename in {"", ".", ".."} or "/" in filename or "\\" in filename:
            raise ValueError(f"Недопустимое имя вложения {filename!r}")

        key = (route_key, message_id, attachment_id)
        self._stage_attachment(key, directory.absolute() / filename, download, temp_suffix, estimated_size)
        destination = self._publish_attachment(key)
        if not self._is_regular_file(destination):
            raise OSError(f"Выгруженный файл недоступен: {destination}")
        return destination

    def _get_attachment(self, key: tuple[str, str, str]) -> tuple[str, int, str | None] | None:
        return self._connection.execute(
            """
            SELECT requested_path, collision_number, staged_path FROM attachment_files
            WHERE route_key = ? AND message_id = ? AND attachment_id = ?
            """,
            key,
        ).fetchone()

    def _stage_attachment(
        self,
        key: tuple[str, str, str],
        requested_path: Path,
        download: Callable[[Path], None],
        temp_suffix: str,
        estimated_size: int,
    ) -> None:
        with self._connection:
            self._connection.execute("BEGIN IMMEDIATE")
            row = self._get_attachment(key)
            if row is not None:
                recorded_path, collision_number, staged_name = row
                requested_path = Path(recorded_path)
                previous_file = Path(staged_name) if staged_name is not None else self._destination(
                    requested_path, collision_number,
                )
                if self._is_regular_file(previous_file):
                    return

            requested_path.parent.mkdir(parents=True, exist_ok=True)
            if shutil.disk_usage(requested_path.parent).free < DISK_RESERVE_BYTES + max(estimated_size, 0):
                raise InsufficientDiskSpace(
                    f"Недостаточно места в {requested_path.parent}. Для выгрузки оставляем резерв 256 МиБ."
                )
            staged_path = None
            try:
                with tempfile.NamedTemporaryFile(
                    dir=requested_path.parent, prefix=".outlook-", suffix=temp_suffix, delete=False
                ) as staged_file:
                    staged_path = Path(staged_file.name)
                download(staged_path)
                if not self._is_regular_file(staged_path):
                    raise OSError(f"Временный файл выгрузки недоступен: {staged_path}")
                if shutil.disk_usage(requested_path.parent).free < DISK_RESERVE_BYTES:
                    raise InsufficientDiskSpace(f"После выгрузки в {requested_path.parent} осталось меньше 256 МиБ.")
                with staged_path.open("r+b") as staged_file:
                    os.fsync(staged_file.fileno())

                self._connection.execute(
                    """
                    INSERT INTO attachment_files
                        (route_key, message_id, attachment_id, requested_path, staged_path)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT (route_key, message_id, attachment_id)
                    DO UPDATE SET staged_path = excluded.staged_path
                    """,
                    (*key, str(requested_path), str(staged_path)),
                )
            except BaseException:
                if staged_path is not None:
                    self._remove_staging_file(staged_path)
                raise

    def _publish_attachment(self, key: tuple[str, str, str]) -> Path:
        while True:
            with self._connection:
                self._connection.execute("BEGIN IMMEDIATE")
                row = self._get_attachment(key)
                if row is None:
                    raise RuntimeError("В журнале отсутствует запись о сохранении вложения")
                requested_path, collision_number, staged_name = row
                original = Path(requested_path)
                destination = self._destination(original, collision_number)
                if staged_name is None:
                    return destination

                staged_path = Path(staged_name)
                if not self._is_regular_file(staged_path):
                    raise OSError(f"Временный файл выгрузки недоступен: {staged_path}")
                try:
                    os.link(staged_path, destination)
                except FileExistsError:
                    if not self._was_published(staged_path, destination):
                        collision_number += 1
                        while os.path.lexists(self._destination(original, collision_number)):
                            collision_number += 1
                        # Record the new path before publishing it so a retry can find its link.
                        self._connection.execute(
                            """
                            UPDATE attachment_files SET collision_number = ?
                            WHERE route_key = ? AND message_id = ? AND attachment_id = ?
                            """,
                            (collision_number, *key),
                        )
                        continue
                except OSError as error:
                    raise OSError(
                        f"Не удалось опубликовать вложение {destination}: {error}. "
                        "Папка выгрузки должна поддерживать жёсткие ссылки."
                    ) from error

                self._connection.execute(
                    """
                    UPDATE attachment_files SET staged_path = NULL
                    WHERE route_key = ? AND message_id = ? AND attachment_id = ?
                    """,
                    key,
                )

            self._remove_staging_file(staged_path)
            return destination

    @staticmethod
    def _is_regular_file(path: Path) -> bool:
        try:
            return stat.S_ISREG(path.lstat().st_mode)
        except (FileNotFoundError, NotADirectoryError):
            return False

    @staticmethod
    def _destination(original: Path, collision_number: int) -> Path:
        if collision_number == 0:
            return original
        return original.with_name(f"{original.stem} ({collision_number}){original.suffix}")

    @staticmethod
    def _was_published(staged_path: Path, destination: Path) -> bool:
        try:
            return not destination.is_symlink() and destination.samefile(staged_path)
        except FileNotFoundError:
            return False

    @staticmethod
    def _remove_staging_file(staged_path: Path) -> None:
        try:
            staged_path.unlink(missing_ok=True)
        except OSError as error:
            LOGGER.warning("Не удалось удалить временный файл %s: %s", staged_path, error)
