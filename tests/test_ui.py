"""Opt-in tests for the real Tk layout and interaction states."""
import os
import tempfile
import time
from pathlib import Path
from unittest.mock import patch
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

    def test_cancel_request_keeps_busy_until_worker_finishes(self):
        self.assertTrue(self.ui.btn_cancel.instate(["disabled"]))
        self.ui.set_busy(True)
        self.assertFalse(self.ui.btn_cancel.instate(["disabled"]))
        self.ui.request_cancel()
        self.assertTrue(self.ui.cancel_event.is_set())
        self.assertTrue(self.ui.is_converting)
        self.assertTrue(self.ui.btn_cancel.instate(["disabled"]))
        self.ui.ui_event_queue.put({"type": "cancelled"})
        self.ui.process_ui_events()
        self.assertFalse(self.ui.is_converting)
        self.assertEqual(self.ui.status_key, "status_cancelled")

    @unittest.skipUnless(os.environ.get("RUN_POWERPOINT_TESTS") == "1", "Requires installed PowerPoint")
    def test_dense_slide_progress_continues_while_minimized(self):
        # Collect destroyed Tk windows on their owning thread before starting the worker.
        import gc
        gc.collect()
        import pythoncom
        import win32com.client
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "dense.pptx"
            pythoncom.CoInitialize()
            try:
                ppt = win32com.client.DispatchEx("PowerPoint.Application")
                try:
                    doc = ppt.Presentations.Add(False)
                    try:
                        slide = doc.Slides.Add(1, 12)
                        for i in range(60):
                            shape = slide.Shapes.AddTextbox(1, (i % 8) * 75, (i // 8) * 45, 72, 40)
                            shape.TextFrame.TextRange.Text = f"항목 {i + 1}"
                        doc.SaveAs(str(path), 24)
                    finally:
                        doc.Close()
                finally:
                    ppt.Quit()
            finally:
                pythoncom.CoUninitialize()
            self.ui.selected_file = str(path)
            self.ui.update_selected_file_label()
            samples = []
            original = self.ui.set_status
            def observe(message, color=None):
                original(message, color)
                samples.append((time.monotonic(), self.ui.lbl_status.cget("text")))
            with patch.object(self.ui, "set_status", side_effect=observe), patch("app.messagebox.askokcancel", return_value=True):
                self.ui.start_conversion()
                self.assertEqual(self.ui.lbl_status.cget("text"), self.ui.tr("opening"))
                self.assertEqual(self.ui.lbl_warning_body.cget("text"), self.ui.tr("minimize_hint"))
                deadline = time.monotonic() + 120
                minimized = False
                while self.ui.is_converting and time.monotonic() < deadline:
                    self.root.update()
                    if not minimized and "도형" in self.ui.lbl_status.cget("text"):
                        self.root.iconify()
                        minimized = True
                    time.sleep(0.02)
                self.assertTrue(minimized)
                self.assertEqual(self.root.state(), "iconic")
                self.root.deiconify()
                self.root.update()
            shapes = [(ts, msg) for ts, msg in samples if "도형" in msg and "/60" in msg]
            self.assertGreater(len(shapes), 2)
            self.assertLess(shapes[0][0], shapes[-1][0])
            self.assertIn(self.ui.tr("status_saving"), [msg for _, msg in samples])
            self.assertEqual(self.ui.result_event["state"], "complete")
            self.assertEqual(self.ui.result_event["converted_count"], 60)
            output = self.ui.result_event["output_path"]
            self.assertTrue(Path(output).is_file())
            with patch("app.os.startfile") as launch:
                self.ui.btn_output.invoke()
                launch.assert_called_once_with(output)
            print(f"Dense slide: {len(shapes)} shape updates over {shapes[-1][0]-shapes[0][0]:.1f}s")


if __name__ == "__main__":
    unittest.main()
