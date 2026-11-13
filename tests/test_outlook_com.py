from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock, patch

from outlook_com import (
    INTERNET_MESSAGE_ID_SCHEMAS,
    MAPI_E_BAD_CHARWIDTH,
    MAPI_E_NO_ACCESS,
    MAPI_E_NO_SUPPORT,
    OutlookError,
    OutlookMailbox,
)


class ComError(Exception):
    def __init__(self, code: int, *, nested: bool = False) -> None:
        super().__init__("COM error")
        signed_code = code - 2**32
        self.hresult = -2147352567 if nested else signed_code
        self.excepinfo = (None, None, "COM error", None, None, signed_code) if nested else None


class MessageIdentityTests(TestCase):
    def setUp(self) -> None:
        com_error_patch = patch("outlook_com._ComError", ComError)
        com_error_patch.start()
        self.addCleanup(com_error_patch.stop)
        self.mailbox = OutlookMailbox.__new__(OutlookMailbox)
        self.mailbox._identity_warning_logged = False
        self.get_property = Mock()
        self.mailbox._mail = Mock(
            return_value=SimpleNamespace(PropertyAccessor=SimpleNamespace(GetProperty=self.get_property))
        )

    def test_falls_back_to_entry_id_when_outlook_denies_message_id(self) -> None:
        self.get_property.side_effect = ComError(MAPI_E_NO_SUPPORT, nested=True)
        with self.assertLogs("outlook_com", level="WARNING") as logs:
            self.assertIsNone(self.mailbox.message_identity("entry", allow_unavailable=True))
            self.assertIsNone(self.mailbox.message_identity("other-entry", allow_unavailable=True))
        self.assertEqual(len(logs.output), 1)
        self.assertEqual(self.get_property.call_count, 2)

    def test_legacy_still_reports_identity_permission_error(self) -> None:
        self.get_property.side_effect = ComError(MAPI_E_NO_ACCESS)
        with self.assertRaises(OutlookError):
            self.mailbox.message_identity("entry")

    def test_tries_other_character_width_when_unicode_is_unsupported(self) -> None:
        self.get_property.side_effect = [ComError(MAPI_E_BAD_CHARWIDTH), "<report@example.com>"]
        self.assertEqual(
            self.mailbox.message_identity("entry", allow_unavailable=True), "<report@example.com>"
        )
        self.assertEqual([call.args[0] for call in self.get_property.call_args_list], list(INTERNET_MESSAGE_ID_SCHEMAS))

    def test_does_not_hide_failure_to_open_message(self) -> None:
        self.mailbox._mail.side_effect = ComError(MAPI_E_NO_ACCESS)
        with self.assertRaises(OutlookError):
            self.mailbox.message_identity("entry", allow_unavailable=True)
