"""Mailbox access through the object model of classic Outlook for Windows."""

from __future__ import annotations

import logging
from datetime import datetime
from email.errors import HeaderParseError
from email.headerregistry import Address
from functools import wraps
from hashlib import sha256
from math import ceil
from pathlib import Path
from time import monotonic, sleep
from typing import Any, Iterable, Iterator

try:
    from pywintypes import com_error as _ComError
except ImportError:
    class _ComError(Exception):
        pass


OL_FOLDER_INBOX = 6
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_JUNK = 23
OL_MAIL = 43
OL_MAIL_ITEM = 0
OL_RULE_RECEIVE = 0
OL_BY_VALUE = 1
OL_YES_NO = 6
OL_MSG_UNICODE = 9
INBOX_NAMES = {"inbox", "входящие"}
PROCESSED_PROPERTY = "Файлы выгружены"
LOGGER = logging.getLogger(__name__)
MAPI_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/"
INTERNET_MESSAGE_ID_SCHEMAS = (
    MAPI_PROPERTY + "0x1035001F", MAPI_PROPERTY + "0x1035001E",
)
SENDER_SMTP_SCHEMAS = (MAPI_PROPERTY + "0x5D01001F", MAPI_PROPERTY + "0x5D01001E")
ADDRESS_ENTRY_SMTP_SCHEMAS = (MAPI_PROPERTY + "0x39FE001E", MAPI_PROPERTY + "0x39FE001F")
EXCHANGE_USER_TYPES = {0, 5}  # Exchange users and remote Exchange users.
MAPI_E_NOT_FOUND = 0x8004010F
MAPI_E_NO_SUPPORT = 0x80040102
MAPI_E_BAD_CHARWIDTH = 0x80040103
MAPI_E_NO_ACCESS = 0x80070005
CATEGORY_COLORS = {
    "Выгружено": 5,
    "Нет файлов": 4,
    "Нет нужных файлов": 4,
    "Ошибка выгрузки": 1,
    "Дубликат": 13,
}


class OutlookError(RuntimeError):
    pass


def _outlook_action(action: str):
    def decorate(method):
        @wraps(method)
        def invoke(*args, **kwargs):
            try:
                return method(*args, **kwargs)
            except _ComError as error:
                info = getattr(error, "excepinfo", None)
                detail = (info and info[2]) or str(error)
                raise OutlookError(f"{action}: {detail}") from error
        return invoke
    return decorate


def _open_session() -> Any:
    try:
        from win32com.client import Dispatch
    except ImportError as error:
        raise OutlookError("Для запуска нужен pywin32 и классический Outlook для Windows.") from error

    try:
        return Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as error:
        raise OutlookError("Не удалось подключиться к классическому Outlook. Проверьте почтовый профиль Windows.") from error


def _category_separator() -> str:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
            separator, _ = winreg.QueryValueEx(key, "sList")
        return separator or ","
    except (ImportError, OSError):
        return ","


def _canonical_folder_path(path: tuple[str, ...]) -> tuple[str, ...]:
    if path and path[0].casefold() in INBOX_NAMES:
        return ("Inbox", *path[1:])
    return path


def _matching_folders(folders: Any, name: str) -> list[Any]:
    folded_name = name.casefold()
    return [folder for folder in folders if folder.Name.casefold() == folded_name]


def _mapi_error_matches(error: _ComError, *expected: int) -> bool:
    info = getattr(error, "excepinfo", None)
    codes = (getattr(error, "hresult", None), info[5] if info and len(info) > 5 else None)
    return any(isinstance(code, int) and (code & 0xFFFFFFFF) in expected for code in codes)


def _mapi_not_found(error: _ComError) -> bool:
    return _mapi_error_matches(error, MAPI_E_NOT_FOUND)


def _text_property(accessor: Any, schemas: tuple[str, ...], *, alternate_width: bool = False) -> str | None:
    for index, schema in enumerate(schemas):
        try:
            value = accessor.GetProperty(schema)
        except _ComError as error:
            if _mapi_not_found(error):
                continue
            if alternate_width and index + 1 < len(schemas) and _mapi_error_matches(error, MAPI_E_BAD_CHARWIDTH):
                continue
            raise
        if value is not None and not isinstance(value, str):
            raise OutlookError("Текстовое свойство Outlook получено в неожиданном формате.")
        if value and value.strip():
            return value
    return None


def _internet_message_id(mail: Any, *, alternate_width: bool = False) -> str | None:
    return _text_property(mail.PropertyAccessor, INTERNET_MESSAGE_ID_SCHEMAS, alternate_width=alternate_width)


def _folder_name(name: str) -> str:
    if not isinstance(name, str) or any(ord(character) < 32 for character in name):
        raise ValueError("Укажите имя новой папки Outlook.")
    name = name.strip()
    if not name or name in {".", ".."} or any(separator in name for separator in ("/", "\\")):
        raise ValueError("Укажите имя одной папки Outlook без разделителей пути.")
    return name


class OutlookMailbox:
    @_outlook_action("Не удалось открыть почтовый профиль Outlook")
    def __init__(self, *, dry_run: bool, manage_categories: bool) -> None:
        self.dry_run = dry_run
        self.manage_categories = manage_categories
        self._separator = _category_separator()
        self._session = _open_session()
        self._inbox = self._session.GetDefaultFolder(OL_FOLDER_INBOX)
        self._store_id = self._inbox.StoreID
        self._identity_warning_logged = False

    def _mail(self, message_id: str) -> Any:
        mail = self._session.GetItemFromID(message_id, self._store_id)
        if mail.Class != OL_MAIL:
            raise OutlookError("Запрошенный элемент Outlook не является письмом.")
        return mail

    def _folder(self, folder_id: str) -> Any:
        return self._session.GetFolderFromID(folder_id, self._store_id)

    def _category_names(self, value: str | None) -> list[str]:
        return [part.strip() for part in (value or "").split(self._separator) if part.strip()]

    def _require_write(self) -> None:
        if self.dry_run:
            raise OutlookError("Пробный режим разрешает только чтение из Outlook.")

    def _matching_children(self, folders: Any, name: str, *, root: bool) -> list[Any]:
        matches = _matching_folders(folders, name)
        if not matches and root and name.casefold() in INBOX_NAMES:
            return [self._inbox]
        return matches

    def _folder_at_path(self, path: tuple[str, ...]) -> Any:
        if path[:1] == ("",):
            folder = self._inbox.Store.GetRootFolder()
            if len(path) < 2 or path[1].casefold() != folder.Name.casefold():
                raise ValueError(f"Укажите папку в основном почтовом ящике Outlook «{folder.Name}».")
            first_child = 2
        elif path and path[0].casefold() in INBOX_NAMES:
            folder = self._inbox
            first_child = 1
        else:
            folder = self._inbox.Store.GetRootFolder()
            first_child = 0

        for index, name in enumerate(path[first_child:], start=first_child):
            matches = self._matching_children(folder.Folders, name, root=(path[:1] == ("",) and index == first_child))
            if len(matches) != 1:
                folder_path = " / ".join(part for part in path[:index + 1] if part)
                if not matches:
                    raise OutlookError(f"В Outlook отсутствует папка «{folder_path}».")
                raise OutlookError(f"В Outlook несколько папок «{folder_path}».")
            folder = matches[0]
        return folder

    @_outlook_action("Не удалось найти папку Outlook")
    def resolve_folder(self, path: tuple[str, ...]) -> str:
        if not path:
            raise OutlookError("Укажите папку Outlook.")
        folder = self._folder_at_path(path)
        if folder.DefaultItemType != OL_MAIL_ITEM:
            raise ValueError("Выберите почтовую папку Outlook для отбора писем.")
        return folder.EntryID

    @_outlook_action("Не удалось проверить папку назначения Outlook")
    def resolve_destination_folder(self, path: tuple[str, ...]) -> str:
        folder = self._folder_at_path(path)
        self._require_allowed_destination(folder)
        if folder.DefaultItemType != OL_MAIL_ITEM:
            raise ValueError("Выберите почтовую папку Outlook для перемещения писем.")
        return folder.EntryID

    def _require_allowed_destination(self, folder: Any) -> None:
        folder_path = folder.FolderPath.casefold()
        store = self._inbox.Store
        for folder_type in (OL_FOLDER_DELETED_ITEMS, OL_FOLDER_OUTBOX, OL_FOLDER_JUNK):
            try:
                protected = store.GetDefaultFolder(folder_type)
            except _ComError as error:
                if not _mapi_not_found(error):
                    raise
                continue
            if protected is not None:
                protected_path = protected.FolderPath.casefold()
                if folder_path == protected_path or folder_path.startswith(protected_path + "\\"):
                    raise ValueError("Для переноса выберите папку вне «Удалённых», «Исходящих» и «Нежелательной почты».")

    def _destination_folder(self, path: tuple[str, ...], *, create: bool) -> Any:
        if not path or (path[:1] == ("",) and len(path) < 3):
            raise ValueError("Укажите почтовую папку назначения Outlook.")
        if len(path) == 1 and path[0].casefold() in INBOX_NAMES:
            raise ValueError("В качестве папки назначения выберите папку вне системных «Входящих».")

        parent = self._folder_at_path(path[:-1])
        siblings = parent.Folders
        name = _folder_name(path[-1])
        matches = self._matching_children(siblings, name, root=(path[:1] == ("",) and len(path) == 3))
        if not matches:
            self._require_allowed_destination(parent)
            if not create:
                return None
            matches = [siblings.Add(name, OL_FOLDER_INBOX)]
        if len(matches) != 1:
            raise OutlookError(f"В Outlook несколько папок «{' / '.join(part for part in path if part)}».")

        folder = matches[0]
        self._require_allowed_destination(folder)
        if folder.EntryID == self._inbox.EntryID:
            raise ValueError("В качестве папки назначения выберите папку вне системных «Входящих».")
        if folder.DefaultItemType != OL_MAIL_ITEM:
            raise ValueError("Выберите почтовую папку Outlook для перемещения писем.")
        return folder

    @_outlook_action("Не удалось проверить папку назначения Outlook")
    def validate_destination_folder(self, path: tuple[str, ...]) -> str | None:
        folder = self._destination_folder(path, create=False)
        return folder.EntryID if folder is not None else None

    @_outlook_action("Не удалось подготовить папку назначения Outlook")
    def ensure_destination_folder(self, path: tuple[str, ...]) -> str:
        self._require_write()
        return self._destination_folder(path, create=True).EntryID

    @_outlook_action("Не удалось получить список папок Outlook")
    def list_folders(self, path: tuple[str, ...]) -> list[dict[str, Any]]:
        prefix = _canonical_folder_path(path)
        children = []
        inbox_id = self._inbox.EntryID if not path else None
        for folder in self._folder_at_path(path).Folders:
            name = folder.Name
            if not path and folder.EntryID == inbox_id:
                segment = "Inbox"
            elif not path and name.casefold() in INBOX_NAMES:
                # These names already address the system Inbox in existing configurations.
                continue
            else:
                segment = name
            children.append({"name": name, "path": [*prefix, segment]})
        return sorted(children, key=lambda child: child["name"].casefold())

    @_outlook_action("Не удалось создать папку Outlook")
    def create_folder(self, parent: tuple[str, ...], name: str) -> dict[str, Any]:
        self._require_write()
        name = _folder_name(name)
        if not parent and name.casefold() in INBOX_NAMES:
            raise ValueError("Это имя используется для системной папки «Входящие».")

        siblings = self._folder_at_path(parent).Folders
        if _matching_folders(siblings, name):
            raise ValueError(f"Папка «{name}» уже есть на этом уровне Outlook.")
        folder = siblings.Add(name, OL_FOLDER_INBOX)
        return {"name": folder.Name, "path": [*_canonical_folder_path(parent), folder.Name]}

    @_outlook_action("Не удалось создать правило Outlook")
    def create_incoming_rule(
        self, path: tuple[str, ...], sender: str, subject_contains: str,
    ) -> dict[str, Any]:
        self._require_write()
        sender = sender.strip() if isinstance(sender, str) else ""
        if not sender or any(ord(character) < 32 for character in sender):
            raise ValueError("Укажите один полный почтовый адрес отправителя.")
        try:
            address = Address(addr_spec=sender)
        except (ValueError, HeaderParseError, IndexError, AttributeError) as error:
            raise ValueError("Укажите один полный почтовый адрес отправителя.") from error
        if not address.username or not address.domain or address.addr_spec != sender:
            raise ValueError("Укажите один полный почтовый адрес отправителя.")
        if not isinstance(subject_contains, str) or any(ord(character) < 32 for character in subject_contains):
            raise ValueError("Укажите фрагмент темы письма одной строкой.")
        subject_contains = subject_contains.strip()
        if not subject_contains:
            raise ValueError("Укажите фрагмент темы письма.")

        if not path:
            raise ValueError("Выберите папку для входящих писем правила.")
        folder_id = self.resolve_destination_folder(path)
        if folder_id == self._inbox.EntryID:
            raise ValueError("Выберите папку, в которую Outlook будет переносить письма из «Входящих».")
        folder = self._folder(folder_id)
        identity = "\0".join((address.username, address.domain.casefold(), subject_contains, folder_id))
        name = "Outlook Exporter " + sha256(identity.encode("utf-8")).hexdigest()[:16]
        rules = self._inbox.Store.GetRules()
        for index in range(1, rules.Count + 1):
            if rules.Item(index).Name.casefold() == name.casefold():
                return {"name": name, "created": False}

        rule = rules.Create(name, OL_RULE_RECEIVE)
        from_condition = rule.Conditions.From
        from_condition.Recipients.Add(sender)
        if not from_condition.Recipients.ResolveAll():
            raise OutlookError("Outlook не смог найти отправителя по этому адресу.")
        from_condition.Enabled = True

        subject_condition = rule.Conditions.Subject
        # pywin32 marshals this tuple as the SAFEARRAY(VARIANT) expected by Outlook.
        subject_condition.Text = (subject_contains,)
        subject_condition.Enabled = True

        move_action = rule.Actions.MoveToFolder
        move_action.Folder = folder
        move_action.Enabled = True

        rule.ExecutionOrder = rules.Count
        rule.Enabled = True
        rules.Save(False)
        return {"name": name, "created": True}

    @_outlook_action("Не удалось прочитать письма Outlook")
    def iter_messages(
        self, folder_id: str, *, read_legacy_status: bool = True, received_since: datetime | None = None,
    ) -> Iterator[dict[str, Any]]:
        items = self._folder(folder_id).Items
        sorted_by_received = False
        if received_since is not None:
            try:
                items.Sort("[ReceivedTime]", True)
                sorted_by_received = True
            except _ComError:
                LOGGER.debug("Outlook не смог отсортировать письма по дате получения", exc_info=True)
        messages = []
        for mail in items:
            if mail.Class != OL_MAIL:
                continue
            entry_id = mail.EntryID
            if not entry_id:
                continue
            # Outlook supplies local wall time as a COM DATE with no zone.
            received = mail.ReceivedTime.replace(tzinfo=None)
            if received_since is not None and received < received_since:
                if sorted_by_received:
                    break
                continue
            message = {
                "id": entry_id,
                "receivedDateTime": received,
                "subject": mail.Subject or "",
                "unread": bool(mail.UnRead),
            }
            if read_legacy_status:
                self._read_legacy_status(mail, message)
            messages.append(message)
        # Processing can move items out of the folder. Do not enumerate a changing collection.
        return iter(messages)

    def _read_legacy_status(self, mail: Any, message: dict[str, Any]) -> None:
        processed = mail.UserProperties.Find(PROCESSED_PROPERTY)
        message["categories"] = self._category_names(mail.Categories)
        message["_processed"] = processed is not None and bool(processed.Value)

    @_outlook_action("Не удалось прочитать адрес отправителя Outlook")
    def message_sender(self, message_id: str) -> str:
        mail = self._mail(message_id)
        if mail.SenderEmailType != "EX":
            return (mail.SenderEmailAddress or "").strip()

        smtp = _text_property(mail.PropertyAccessor, SENDER_SMTP_SCHEMAS)
        if smtp and "@" in smtp:
            return smtp.strip()

        sender = mail.Sender
        if sender is None:
            return ""
        if sender.AddressEntryUserType in EXCHANGE_USER_TYPES:
            user = sender.GetExchangeUser()
            if user is not None and user.PrimarySmtpAddress:
                return user.PrimarySmtpAddress.strip()
        smtp = _text_property(sender.PropertyAccessor, ADDRESS_ENTRY_SMTP_SCHEMAS)
        return (smtp or "").strip()

    @_outlook_action("Не удалось прочитать идентификатор письма Outlook")
    def message_identity(self, message_id: str, *, allow_unavailable: bool = False) -> str | None:
        mail = self._mail(message_id)
        try:
            return _internet_message_id(mail, alternate_width=allow_unavailable)
        except _ComError as error:
            if not allow_unavailable or not _mapi_error_matches(
                error, MAPI_E_NO_SUPPORT, MAPI_E_NO_ACCESS, MAPI_E_BAD_CHARWIDTH,
            ):
                raise
            if not self._identity_warning_logged:
                LOGGER.warning("Outlook не предоставил Message-ID. Для таких писем журнал использует EntryID Outlook.")
                self._identity_warning_logged = True
            return None

    @_outlook_action("Не удалось проверить копии письма в папке назначения Outlook")
    def find_destination_matches(
        self, folder_id: str, identity: str, *, read_legacy_status: bool = True,
    ) -> list[dict[str, Any]]:
        escaped_identity = identity.replace("'", "''")
        conditions = [f"\"{schema}\" = '{escaped_identity}'" for schema in INTERNET_MESSAGE_ID_SCHEMAS]
        items = self._folder(folder_id).Items
        try:
            matches = items.Restrict("@SQL=" + " OR ".join(conditions))
        except _ComError:
            LOGGER.debug("Outlook не смог выполнить поиск копий, используется просмотр папки", exc_info=True)
            matches = items
        result = []
        for mail in matches:
            # The store may compare strings case-insensitively when applying the restriction.
            if mail.Class != OL_MAIL or _internet_message_id(mail, alternate_width=not read_legacy_status) != identity:
                continue
            match = {"id": mail.EntryID}
            if read_legacy_status:
                self._read_legacy_status(mail, match)
            result.append(match)
        return result

    @_outlook_action("Не удалось прочитать вложения Outlook")
    def iter_attachments(self, message_id: str) -> Iterator[dict[str, Any]]:
        attachments = self._mail(message_id).Attachments
        files = []
        for index in range(1, attachments.Count + 1):
            attachment = attachments.Item(index)
            files.append({
                "id": str(index),
                "name": attachment.FileName,
                "is_file": attachment.Type == OL_BY_VALUE,
                "size": int(attachment.Size),
            })
        return iter(files)

    @_outlook_action("Не удалось сохранить вложение Outlook")
    def save_attachment(self, message_id: str, attachment_id: str, output: Path) -> None:
        self._require_write()
        attachment = self._mail(message_id).Attachments.Item(int(attachment_id))
        if attachment.Type != OL_BY_VALUE:
            raise OutlookError("Запрошенное вложение Outlook не является файлом.")
        attachment.SaveAsFile(str(output.resolve()))

    @_outlook_action("Не удалось сохранить письмо Outlook")
    def save_message(self, message_id: str, output: Path) -> None:
        self._require_write()
        if output.suffix.casefold() != ".msg":
            raise ValueError("Для сохранения письма укажите путь с расширением .msg.")
        self._mail(message_id).SaveAs(str(output.resolve()), OL_MSG_UNICODE)

    @_outlook_action("Не удалось прочитать размер письма Outlook")
    def message_size(self, message_id: str) -> int:
        return int(self._mail(message_id).Size)

    @_outlook_action("Не удалось настроить категории Outlook")
    def ensure_categories(self) -> None:
        if not self.manage_categories or self.dry_run:
            return
        categories = self._session.Categories
        existing = {category.Name.casefold(): category for category in categories}
        for name, color in CATEGORY_COLORS.items():
            category = existing.get(name.casefold())
            if category is None:
                categories.Add(name, color)
            elif category.Color != color:
                category.Color = color

    @staticmethod
    def _mark_read_if_allowed(mail: Any) -> bool:
        try:
            if mail.UnRead and mail.ReadReceiptRequested is False:
                mail.UnRead = False
                return True
        except (AttributeError, _ComError):
            # A provider may not expose the receipt flag or allow changing read state.
            pass
        return False

    @_outlook_action("Не удалось отметить письмо прочитанным в Outlook")
    def mark_read(self, message_id: str) -> None:
        self._require_write()
        mail = self._mail(message_id)
        if self._mark_read_if_allowed(mail):
            mail.Save()

    @_outlook_action("Не удалось записать результат в Outlook")
    def set_result(self, message_id: str, status: str, *, processed: bool, mark_read: bool = False) -> None:
        self._require_write()
        if status not in CATEGORY_COLORS:
            raise ValueError(f"Неизвестный статус выгрузки: {status}")
        mail = self._mail(message_id)
        # A user may have updated categories since the scan began.
        current_categories = self._category_names(mail.Categories)
        status_names = {name.casefold() for name in CATEGORY_COLORS}
        updated = [name for name in current_categories if name.casefold() not in status_names]
        updated.append(status)

        mail.Categories = f"{self._separator} ".join(updated)
        if processed:
            properties = mail.UserProperties
            flag = properties.Find(PROCESSED_PROPERTY)
            if flag is None:
                flag = properties.Add(PROCESSED_PROPERTY, OL_YES_NO, True)
            flag.Value = True
        if mark_read:
            self._mark_read_if_allowed(mail)
        mail.Save()

    @staticmethod
    def is_processed(message: dict[str, Any]) -> bool:
        return bool(message.get("_processed"))

    @_outlook_action("Не удалось проверить расположение письма в Outlook")
    def is_in_folder(self, message_id: str, folder_id: str) -> bool:
        try:
            return self._mail(message_id).Parent.EntryID == folder_id
        except _ComError as error:
            if _mapi_not_found(error):
                return False
            raise

    @_outlook_action("Не удалось переместить письмо в Outlook")
    def move_message(self, message_id: str, folder_id: str, *, expected_source: str | None = None) -> str:
        self._require_write()
        destination = self._folder(folder_id)
        mail = self._mail(message_id)
        if expected_source is not None and mail.Parent.EntryID != expected_source:
            raise OutlookError("Письмо уже находится вне исходной папки. Перемещение пропущено.")
        moved = mail.Move(destination)
        moved_id = moved.EntryID if moved is not None else None
        if not moved_id:
            raise OutlookError("Outlook не вернул идентификатор письма после перемещения.")
        return moved_id

    def wait_for_changes(self, folder_ids: Iterable[str], timeout: int) -> None:
        """Wake on Outlook arrivals, with a timed scan for missed or unavailable events."""
        deadline = monotonic() + timeout
        subscriptions = []
        arrived = False

        class FolderEvents:
            def OnItemAdd(self, _item):
                nonlocal arrived
                arrived = True

        try:
            import pythoncom
            import win32event
            from win32com.client import WithEvents

            for folder_id in dict.fromkeys(folder_ids):
                try:
                    items = self._folder(folder_id).Items
                    subscriptions.append((items, WithEvents(items, FolderEvents)))
                except Exception:
                    # WithEvents may fail while generating or loading the Outlook type library.
                    LOGGER.debug("Не удалось подписаться на события папки Outlook", exc_info=True)
            if subscriptions:
                while not arrived:
                    remaining = deadline - monotonic()
                    if remaining <= 0:
                        return
                    # Keep the wait interruptible by Ctrl+C on Windows.
                    win32event.MsgWaitForMultipleObjects((), False, min(250, ceil(remaining * 1000)), win32event.QS_ALLINPUT)
                    if pythoncom.PumpWaitingMessages():
                        raise KeyboardInterrupt
                return
        except Exception:
            LOGGER.debug("События Outlook недоступны, используется опрос по таймеру", exc_info=True)
        finally:
            for _, subscription in subscriptions:
                try:
                    subscription.close()
                except Exception:
                    LOGGER.debug("Не удалось отключить события папки Outlook", exc_info=True)
        sleep(max(0, deadline - monotonic()))
