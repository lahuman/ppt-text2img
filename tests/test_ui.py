"""Opt-in tests for the real Tk layout and interaction states."""
import os
import tkinter as tk
import unittest

import app


@unittest.skipUnless(os.environ.get("RUN_UI_TESTS") == "1", "Requires a desktop session")
class InterfaceTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.ui = app.App(self.root)
        self.root.update()

    def tearDown(self):
        self.root.destroy()

    def test_languages_and_long_path_fit_workspace(self):
        for language in ("ko", "en"):
            self.ui.lang_code = language
            self.ui.apply_language()
            self.ui.selected_file = "C:/Presentations/" + "Long presentation name " * 12 + ".pptx"
            self.ui.update_selected_file_label()
            self.root.update()
            self.assertEqual(self.root.title(), app.LANG[language]["app_title"])
            self.assertLessEqual(self.ui.lbl_file_value.winfo_reqheight(), 65)
            for widget in (self.ui.lbl_title, self.ui.file_panel, self.ui.progress_panel,
                           self.ui.btn_convert, self.ui.lbl_requirements):
                self.assertGreater(widget.winfo_width(), 100)
                self.assertLessEqual(widget.winfo_y() + widget.winfo_height(),
                                     self.ui.container.winfo_height())

    def test_busy_state_and_progress(self):
        self.assertTrue(self.ui.btn_convert.instate(["disabled"]))
        self.ui.selected_file = "C:/Presentations/talk.pptx"
        self.ui.update_selected_file_label()
        self.assertFalse(self.ui.btn_convert.instate(["disabled"]))
        self.ui.set_busy(True)
        self.assertTrue(self.ui.btn_select.instate(["disabled"]))
        self.assertTrue(self.ui.btn_convert.instate(["disabled"]))
        self.ui.update_progress(3, 5, "Processing slide 3/5")
        self.assertEqual(self.ui.lbl_percent.cget("text"), "60%")
        self.ui.set_busy(False)
        self.assertFalse(self.ui.btn_select.instate(["disabled"]))

    def test_small_window_keeps_actions_reachable_by_scrolling(self):
        self.root.geometry("940x650")
        self.root.update()
        self.assertTrue(self.ui.scrollbar.winfo_ismapped())
        self.ui.content_canvas.yview_moveto(1)
        self.root.update()
        self.assertGreater(self.ui.content_canvas.yview()[0], 0)


if __name__ == "__main__":
    unittest.main()
