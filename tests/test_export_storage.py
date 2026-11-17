import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from outlook_storage import Storage


class WindowsPublicationTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.output = Path(directory.name) / "output"
        self.output.mkdir()
        self.database = self.output.parent / "state.sqlite3"
        self.download = Mock(side_effect=lambda file: file.write_bytes(b"attachment bytes"))
        self.real_rename = os.rename

    def rename_without_overwrite(self, source, destination):
        if os.path.lexists(destination):
            raise FileExistsError(destination)
        self.real_rename(source, destination)

    def save(self):
        storage = Storage(self.database)
        try:
            return storage.save_attachment("rule", "mail", "1", self.output, "report.csv", self.download)
        finally:
            storage.close()

    def test_publication_does_not_require_hard_links_or_overwrite_existing_file(self):
        original = self.output / "report.csv"
        original.write_bytes(b"existing bytes")
        with (
            patch("outlook_storage.USE_NOCLOBBER_RENAME", True),
            patch("outlook_storage.os.rename", side_effect=self.rename_without_overwrite),
            patch("outlook_storage.os.link", side_effect=OSError("Hard links unsupported")) as link,
        ):
            result = self.save()

        link.assert_not_called()
        self.assertEqual(result.name, "report (1).csv")
        self.assertEqual(original.read_bytes(), b"existing bytes")
        self.assertEqual(result.read_bytes(), b"attachment bytes")
        self.assertEqual(set(self.output.iterdir()), {original, result})

    def test_recovers_when_interrupted_after_rename(self):
        def interrupted_rename(source, destination):
            self.rename_without_overwrite(source, destination)
            raise KeyboardInterrupt

        with patch("outlook_storage.USE_NOCLOBBER_RENAME", True):
            with patch("outlook_storage.os.rename", side_effect=interrupted_rename):
                with self.assertRaises(KeyboardInterrupt):
                    self.save()
            with patch("outlook_storage.os.rename", side_effect=self.rename_without_overwrite) as rename:
                result = self.save()

        rename.assert_not_called()
        self.download.assert_called_once()
        self.assertEqual(result.read_bytes(), b"attachment bytes")
        self.assertEqual(list(self.output.iterdir()), [result])

    def test_recovery_does_not_accept_different_bytes_at_destination(self):
        def interrupted_rename(source, destination):
            self.rename_without_overwrite(source, destination)
            raise KeyboardInterrupt

        with patch("outlook_storage.USE_NOCLOBBER_RENAME", True):
            with patch("outlook_storage.os.rename", side_effect=interrupted_rename):
                with self.assertRaises(KeyboardInterrupt):
                    self.save()
            original = self.output / "report.csv"
            original.write_bytes(b"unrelated bytes")
            with patch("outlook_storage.os.rename", side_effect=self.rename_without_overwrite):
                result = self.save()

        self.assertEqual(self.download.call_count, 2)
        self.assertEqual(original.read_bytes(), b"unrelated bytes")
        self.assertEqual(result.name, "report (1).csv")
        self.assertEqual(result.read_bytes(), b"attachment bytes")
