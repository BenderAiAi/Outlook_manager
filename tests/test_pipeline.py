import json
import tempfile
import unittest
from datetime import datetime, timedelta
from http.client import HTTPConnection
from pathlib import Path
from threading import Thread
from unittest.mock import patch

from outlook_config import load_settings, validate_settings, write_settings
from outlook_exporter import scan
from outlook_setup import SetupServer
from outlook_storage import DISK_RESERVE_BYTES, Storage


class FakeMailbox:
    def __init__(self):
        self.folders = {"Inbox": {}}
        self.moves = []
        self.denied = set()
        self.save_calls = []
        self.failed_downloads = set()

    def add(self, key, *, folder="Inbox", subject="Report", sender="reports@example.com", files=(),
            received=None, identity=None):
        self.folders.setdefault(folder, {})[key] = {
            "id": key,
            "receivedDateTime": received or datetime.now() + timedelta(seconds=5),
            "subject": subject,
            "unread": True,
            "sender": sender,
            "files": list(files),
            "identity": identity,
        }

    def resolve_folder(self, path):
        return "/".join(path)

    def validate_destination_folder(self, path):
        key = self.resolve_folder(path)
        return key if key in self.folders else None

    def ensure_destination_folder(self, path):
        key = self.resolve_folder(path)
        self.folders.setdefault(key, {})
        return key

    def iter_messages(self, folder_id, *, read_legacy_status=True, received_since=None):
        return iter(message.copy() for message in self.folders[folder_id].values()
                    if received_since is None or message["receivedDateTime"] >= received_since)

    def _message(self, message_id):
        return next(folder[message_id] for folder in self.folders.values() if message_id in folder)

    def message_sender(self, message_id):
        return self._message(message_id)["sender"]

    def message_identity(self, message_id, *, allow_unavailable=False):
        return self._message(message_id)["identity"]

    def find_destination_matches(self, folder_id, identity, *, read_legacy_status=True):
        return [message.copy() for message in self.folders[folder_id].values() if message["identity"] == identity]

    def is_in_folder(self, message_id, folder_id):
        return message_id in self.folders[folder_id]

    def move_message(self, message_id, folder_id, *, expected_source=None):
        if folder_id in self.denied:
            raise RuntimeError("Access denied")
        message = self.folders[expected_source].pop(message_id)
        new_id = message_id + "/moved"
        message["id"] = new_id
        self.folders[folder_id][new_id] = message
        self.moves.append((expected_source, folder_id))
        return new_id

    def iter_attachments(self, message_id):
        return iter({"id": str(index), "name": name, "is_file": True, "size": 5}
                    for index, name in enumerate(self._message(message_id)["files"], start=1))

    def save_attachment(self, message_id, attachment_id, output):
        self.save_calls.append((message_id, attachment_id))
        if attachment_id in self.failed_downloads:
            raise OSError("Unable to save attachment")
        output.write_bytes(b"file\n")

    def message_size(self, message_id):
        return 8

    def save_message(self, message_id, output):
        self.save_calls.append((message_id, "msg"))
        output.write_bytes(b"message\n")

    def mark_read(self, message_id):
        self._message(message_id)["unread"] = False


class PipelineTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.config = Path(directory.name) / "config.toml"
        self.data = {
            "manage_categories": False,
            "routes": [{
                "pending": ["Inbox", "Pending"],
                "processed": ["Inbox", "Processed"],
                "error": ["Inbox", "Error"],
                "subject_exact": "Report",
                "output_dir": "exported",
            }],
        }
        self.route = self.data["routes"][0]
        self.mailbox = FakeMailbox()

    def run_scan(self, *, dry_run=False):
        settings = validate_settings(self.data, self.config)
        if dry_run:
            return scan(self.mailbox, settings, None)
        storage = Storage(settings.state_db)
        try:
            return scan(self.mailbox, settings, storage)
        finally:
            storage.close()

    def test_exact_match_without_attachments_moves_through_queue(self):
        self.route["sender"] = "reports@example.com"
        self.mailbox.add("match", subject="REPORT")
        self.mailbox.add("wrong-subject", subject="Report supplement")
        self.mailbox.add("wrong-sender", sender="someone@example.com")
        self.assertEqual(self.run_scan(), 0)
        self.assertEqual(self.mailbox.moves, [("Inbox", "Inbox/Pending"), ("Inbox/Pending", "Inbox/Processed")])
        self.assertEqual(set(self.mailbox.folders["Inbox"]), {"wrong-subject", "wrong-sender"})
        self.assertFalse(self.mailbox.save_calls)

    def test_first_matching_rule_wins_in_shared_source(self):
        self.data["routes"].append({**self.route, "pending": ["Inbox", "Second"]})
        self.mailbox.add("match", files=("report.csv",))
        self.assertEqual(self.run_scan(), 0)
        self.assertEqual(self.mailbox.moves, [("Inbox", "Inbox/Pending"), ("Inbox/Pending", "Inbox/Processed")])
        self.assertEqual(self.mailbox.save_calls, [("match/moved", "1")])
        self.assertFalse(self.mailbox.folders["Inbox/Second"])

    def test_existing_queue_is_processed_regardless_of_incoming_filter(self):
        yesterday = datetime.now() - timedelta(days=1)
        self.mailbox.add("queued", folder="Inbox/Pending", subject="Other", received=yesterday, files=("old.csv",))
        self.mailbox.add("inbox-old", received=yesterday)
        self.assertEqual(self.run_scan(), 0)
        self.assertIn("inbox-old", self.mailbox.folders["Inbox"])
        self.assertIn("queued/moved", self.mailbox.folders["Inbox/Processed"])
        self.assertEqual(self.mailbox.save_calls, [("queued", "1")])

    def test_dry_run_only_reads_and_live_saves_msg_and_attachment(self):
        self.route["save_message"] = True
        self.mailbox.add("match", files=("report.csv",))
        self.assertEqual(self.run_scan(dry_run=True), 0)
        self.assertEqual(set(self.mailbox.folders), {"Inbox"})
        self.assertFalse(self.mailbox.moves)
        self.assertFalse(self.mailbox.save_calls)
        self.assertEqual(list(self.config.parent.iterdir()), [])
        self.assertEqual(self.run_scan(), 0)
        files = list((self.config.parent / "exported").iterdir())
        self.assertEqual({path.suffix: path.read_bytes() for path in files}, {".csv": b"file\n", ".msg": b"message\n"})

    def test_duplicate_goes_to_error_without_saving_copy(self):
        self.mailbox.add("first", identity="<shared@example.com>", files=("report.csv",))
        self.assertEqual(self.run_scan(), 0)
        self.mailbox.add("copy", identity="<shared@example.com>", files=("report.csv",))
        self.run_scan()
        self.assertIn("copy/moved/moved", self.mailbox.folders["Inbox/Error"])
        self.assertEqual(len(self.mailbox.save_calls), 1)

    def test_failed_error_move_retries_without_duplicate_save(self):
        self.mailbox.add("match", files=("report.csv",))
        self.mailbox.denied.update({"Inbox/Processed", "Inbox/Error"})
        self.assertGreater(self.run_scan(), 0)
        self.assertIn("match/moved", self.mailbox.folders["Inbox/Pending"])
        self.mailbox.denied.remove("Inbox/Error")
        self.run_scan()
        self.assertIn("match/moved/moved", self.mailbox.folders["Inbox/Error"])
        self.assertEqual(len(self.mailbox.save_calls), 1)

    def test_intake_error_retry_takes_priority_over_changed_filter(self):
        self.route["subject_exact"] = "Other"
        self.data["routes"].append({**self.route, "subject_exact": "Report", "pending": ["Inbox", "Second"]})
        self.mailbox.add("match", files=("report.csv",))
        self.mailbox.denied.update({"Inbox/Second", "Inbox/Error"})
        self.assertGreater(self.run_scan(), 0)
        self.assertIn("match", self.mailbox.folders["Inbox"])
        self.route["subject_exact"] = "Report"
        self.mailbox.denied.clear()
        self.assertEqual(self.run_scan(), 0)
        self.assertIn("match/moved", self.mailbox.folders["Inbox/Error"])
        self.assertFalse(self.mailbox.save_calls)

    def test_intake_error_retry_survives_turning_off_historical_mail(self):
        self.route["include_existing"] = True
        self.mailbox.add("old", received=datetime.now() - timedelta(days=1))
        self.mailbox.denied.update({"Inbox/Pending", "Inbox/Error"})
        self.assertGreater(self.run_scan(), 0)
        self.assertIn("old", self.mailbox.folders["Inbox"])

        self.route["include_existing"] = False
        self.mailbox.denied.clear()
        self.assertEqual(self.run_scan(), 0)
        self.assertIn("old/moved", self.mailbox.folders["Inbox/Error"])

    def test_cutoff_survives_restart(self):
        self.assertEqual(self.run_scan(), 0)
        self.mailbox.add("late-sync", received=datetime.now() + timedelta(seconds=1))
        self.assertEqual(self.run_scan(), 0)
        self.assertIn("late-sync/moved/moved", self.mailbox.folders["Inbox/Processed"])

    def test_contains_with_attachment_filter_and_historical_mail(self):
        self.route.pop("subject_exact")
        self.route.update(subject_contains="report", attachment_filter="present", include_existing=True)
        yesterday = datetime.now() - timedelta(days=1)
        self.mailbox.add("old-file", received=yesterday, subject="Daily REPORT", files=("report.csv",))
        self.mailbox.add("old-no-file", received=yesterday, subject="Daily Report")
        self.assertEqual(self.run_scan(), 0)
        self.assertIn("old-file/moved/moved", self.mailbox.folders["Inbox/Processed"])
        self.assertIn("old-no-file", self.mailbox.folders["Inbox"])

    def test_save_failure_moves_to_error_and_records_nonzero_result(self):
        self.mailbox.add("broken-file", files=("report.csv",))
        self.mailbox.failed_downloads.add("1")
        self.assertGreater(self.run_scan(), 0)
        self.assertIn("broken-file/moved/moved", self.mailbox.folders["Inbox/Error"])
        self.assertFalse(list((self.config.parent / "exported").iterdir()))

    def test_low_disk_retains_mail_in_queue_until_next_scan(self):
        self.mailbox.add("match", files=("report.csv",))
        with patch("outlook_storage.shutil.disk_usage") as disk_usage:
            disk_usage.return_value.free = DISK_RESERVE_BYTES
            self.assertGreater(self.run_scan(), 0)
            self.assertIn("match/moved", self.mailbox.folders["Inbox/Pending"])
            self.assertFalse(self.mailbox.folders["Inbox/Error"])
            self.assertFalse(self.mailbox.save_calls)
            disk_usage.return_value.free = DISK_RESERVE_BYTES + 1024
            self.assertEqual(self.run_scan(), 0)
        self.assertIn("match/moved/moved", self.mailbox.folders["Inbox/Processed"])
        self.assertEqual(self.mailbox.save_calls, [("match/moved", "1")])

    def test_toml_round_trip_for_pasted_paths_and_legacy_route(self):
        self.route["pending"] = r"\\Почтовый ящик\Входящие\Очередь"
        self.route["output_dir"] = '"./отчёты с пробелами"'
        self.data["routes"].append({"source": ["Old"], "output_dir": "old", "move_to": ["Old processed"]})
        write_settings(self.data, self.config)
        pipeline, legacy = load_settings(self.config).routes
        self.assertEqual(pipeline.pending, ("", "Почтовый ящик", "Входящие", "Очередь"))
        self.assertEqual(pipeline.output_dir.name, "отчёты с пробелами")
        self.assertEqual(legacy.move_to, ("Old processed",))
        self.assertIsNone(legacy.pending)

    def test_wizard_previews_and_saves_without_moving_mail(self):
        self.mailbox.add("matched")
        self.mailbox.add("other", subject="Other")
        with SetupServer(self.config, 0) as server, patch("outlook_setup.open_mailbox", return_value=self.mailbox):
            thread = Thread(target=server.serve_forever)
            thread.start()
            client = HTTPConnection("127.0.0.1", server.server_port, timeout=5)
            try:
                client.request("GET", "/api/settings")
                response = client.getresponse()
                self.assertEqual(response.status, 200)
                token = json.load(response)["csrf_token"]
                headers = {"Content-Type": "application/json", "Origin": server.url, "X-CSRF-Token": token}
                client.request("POST", "/api/preview", json.dumps({"settings": self.data, "route_index": 0}), headers)
                response = client.getresponse()
                preview = json.load(response)
                self.assertEqual(response.status, 200, preview)
                self.assertEqual((preview["scanned"], preview["matched"], preview["pending"]), (2, 1, 0))
                client.request("POST", "/api/settings", json.dumps({"settings": self.data}), headers)
                response = client.getresponse()
                self.assertEqual(response.status, 200, json.load(response))
            finally:
                client.close()
                server.shutdown()
                thread.join()
        self.assertEqual(load_settings(self.config).routes[0].subject_exact, "Report")
        self.assertEqual(set(self.mailbox.folders), {"Inbox"})
        self.assertFalse(self.mailbox.moves)


if __name__ == "__main__":
    unittest.main()
