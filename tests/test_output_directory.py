import ntpath
import sys
from pathlib import PureWindowsPath
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock, patch

import outlook_config
import outlook_setup


class WindowsPath(PureWindowsPath):
    def expanduser(self):
        return self

    def resolve(self):
        return self

    def is_dir(self):
        return False

    def is_file(self):
        return False


class OutputDirectoryTests(TestCase):
    def test_settings_accept_network_paths_from_explorer(self):
        def check_managed_path(path):
            if path.drive != "C:":
                self.fail(f"Export folder incorrectly checked as a settings or journal path: {path}")

        with patch.object(outlook_config, "Path", WindowsPath), patch.object(
            outlook_config, "require_local_path", side_effect=check_managed_path,
        ):
            for selected in (r"Z:\Отчёты\Выгрузка", r"\\server\share\Отчёты"):
                with self.subTest(selected=selected):
                    settings = outlook_config.validate_settings({
                        "routes": [{"source": ["Inbox"], "output_dir": f'"{selected}"'}],
                    }, WindowsPath(r"C:\Exporter\config.toml"))
                    self.assertEqual(settings.routes[0].output_dir, WindowsPath(selected))

    def test_native_dialog_returns_a_network_path_or_cancellation(self):
        owner = Mock()
        askdirectory = Mock(side_effect=["//server/share/Отчёты", "Z:/Отчёты", ""])
        tkinter = SimpleNamespace(Tk=Mock(return_value=owner), TclError=RuntimeError,
                                  filedialog=SimpleNamespace(askdirectory=askdirectory))
        windows = SimpleNamespace(name="nt", path=ntpath)
        with patch.object(outlook_setup, "os", windows), patch.dict(sys.modules, {"tkinter": tkinter}):
            self.assertEqual(outlook_setup.pick_output_directory(), r"\\server\share\Отчёты")
            self.assertEqual(outlook_setup.pick_output_directory(), r"Z:\Отчёты")
            self.assertIsNone(outlook_setup.pick_output_directory())
        askdirectory.assert_called_with(parent=owner, title="Выберите папку для сохранения", mustexist=True)
        self.assertEqual(owner.destroy.call_count, 3)
