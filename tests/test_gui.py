import tkinter as tk
import unittest
from unittest.mock import patch

from autograph_service.gui import AutoGraphApp
from autograph_service.service import (
    DATA_KIND_S2P,
    SOURCE_KIND_DIRECTORY,
)


class GuiSmokeTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as error:  # pragma: no cover - depends on environment
            self.skipTest(str(error))
        self.root.withdraw()
        self.app = AutoGraphApp(master=self.root, runner=self.fake_runner)
        self.root.update()

    def tearDown(self):
        self.app.destroy()
        self.root.destroy()

    def fake_runner(self, config, reporter):
        reporter("fake run")
        return None

    def test_default_visible_fields_show_template_single_txt_mode(self):
        self.assertEqual(
            self.app.get_visible_field_keys(),
            ["input_path", "excel_path", "template_sheet", "charts_sheet", "output_path"],
        )

    def test_plain_txt_single_mode_shows_sheet_name(self):
        self.app.template_mode_var.set(False)
        self.root.update()
        self.assertEqual(
            self.app.get_visible_field_keys(),
            ["input_path", "excel_path", "sheet_name", "output_path"],
        )

    def test_directory_mode_hides_single_file_only_fields(self):
        self.app.source_kind_var.set(SOURCE_KIND_DIRECTORY)
        self.root.update()
        self.assertEqual(
            self.app.get_visible_field_keys(),
            ["input_path", "excel_path", "template_sheet", "charts_sheet"],
        )

    def test_s2p_mode_forces_template_settings(self):
        self.app.data_kind_var.set(DATA_KIND_S2P)
        self.root.update()

        self.assertTrue(self.app.template_mode_var.get())
        self.assertEqual(
            self.app.get_visible_field_keys(),
            ["input_path", "excel_path", "template_sheet", "charts_sheet_template", "use_template_charts", "output_path"],
        )

    def test_validation_shows_error_when_required_paths_missing(self):
        with patch("autograph_service.gui.messagebox.showerror") as mocked:
            self.app.start_run()
        mocked.assert_called_once()


if __name__ == "__main__":
    unittest.main()
