import json
import tempfile
import threading
import unittest
import zipfile
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import MagicMock, PropertyMock, patch

import conversion as app


class LifecycleTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.folder = Path(self.stack.enter_context(tempfile.TemporaryDirectory()))
        self.source = self.folder / "source.pptx"
        self.source.write_bytes(b"original")
        self.com = self.stack.enter_context(patch.object(app, "pythoncom"))
        self.win = self.stack.enter_context(patch.object(app, "win32com"))
        self.stack.enter_context(patch.object(app, "ensure_runtime_requirements"))
        self.ppt = MagicMock(DisplayAlerts=2, AutomationSecurity=1)
        self.ppt.Presentations.__iter__.return_value = iter([])
        self.ppt.Presentations.Count = 0
        self.presentation = MagicMock()
        self.presentation.Slides.Count = 0
        self.presentation.Slides.__iter__.return_value = iter([])
        self.presentation.PageSetup.SlideWidth = 720
        self.presentation.PageSetup.SlideHeight = 540
        self.ppt.Presentations.Open.return_value = self.presentation
        self.win.client.GetActiveObject.return_value = self.ppt
        self.presentation.SaveAs.side_effect = self.write_output

    def write_output(self, path, format_id):
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("ppt/presentation.xml", "<presentation/>")

    def convert(self):
        return app.text_to_image_ppt(str(self.source))

    def test_close_failure_is_reported(self):
        self.presentation.Close.side_effect = RuntimeError("close failed")
        result = self.convert()
        self.assertIn("close", [issue["stage"] for issue in result.issues])
        report = json.loads(Path(result.report_path).read_text(encoding="utf-8"))
        self.assertEqual(report["issues"], result.issues)
        self.com.CoUninitialize.assert_called_once()
        self.assertEqual(self.source.read_bytes(), b"original")

    def test_restore_failure_does_not_skip_other_cleanup(self):
        self.win.client.GetActiveObject.side_effect = RuntimeError("no session")
        self.win.client.DispatchEx.return_value = self.ppt
        alerts = PropertyMock(side_effect=[2, None, RuntimeError("restore failed"), 1, None, None])
        type(self.ppt).DisplayAlerts = alerts
        result = self.convert()
        self.assertIn("restore_alerts", [issue["stage"] for issue in result.issues])
        self.assertEqual(self.ppt.AutomationSecurity, 1)
        self.ppt.Quit.assert_called_once()
        self.com.CoUninitialize.assert_called_once()

    def test_open_failure_preserves_primary_error_and_cleans_up(self):
        self.ppt.Presentations.Open.side_effect = RuntimeError("cannot open")
        with self.assertRaises(app.UserVisibleError) as raised:
            self.convert()
        self.assertEqual(raised.exception.kind, "file_open")
        self.assertEqual(self.ppt.DisplayAlerts, 2)
        self.assertEqual(self.ppt.AutomationSecurity, 1)
        self.com.CoUninitialize.assert_called_once()

    def test_save_failure_closes_document_and_preserves_source(self):
        self.presentation.SaveAs.side_effect = RuntimeError("disk full")
        with self.assertRaises(app.UserVisibleError) as raised:
            self.convert()
        self.assertEqual(raised.exception.kind, "file_save")
        self.assertEqual(self.ppt.DisplayAlerts, 2)
        self.presentation.Close.assert_called_once()
        self.com.CoUninitialize.assert_called_once()
        self.assertEqual(self.source.read_bytes(), b"original")

    def test_alerts_restored_between_open_and_save(self):
        observations = []
        self.ppt.Presentations.__iter__.return_value = iter([MagicMock(FullName="other.pptx")])
        def open_file(*args, **kwargs):
            self.assertEqual(self.ppt.DisplayAlerts, 1)
            return self.presentation

        def save_file(path, fmt):
            self.assertEqual(self.ppt.DisplayAlerts, 1)
            self.write_output(path, fmt)

        self.ppt.Presentations.Open.side_effect = open_file
        self.presentation.SaveAs.side_effect = save_file
        app.text_to_image_ppt(str(self.source), lambda *args: observations.append(self.ppt.DisplayAlerts))
        self.assertTrue(observations)
        self.assertEqual(set(observations), {2})
        self.assertEqual(self.ppt.DisplayAlerts, 2)

    def test_existing_session_is_never_quit(self):
        result = self.convert()
        self.assertEqual(result.state, "complete")
        self.ppt.Quit.assert_not_called()
        self.assertTrue(Path(result.output_path).is_file())
        self.assertFalse(list(self.folder.glob(".ppt-pixel-pending-*")))

    def test_cancel_before_start_never_initializes_com(self):
        cancelled = threading.Event()
        cancelled.set()
        with self.assertRaises(app.ConversionCancelled):
            app.text_to_image_ppt(str(self.source), cancel_event=cancelled)
        self.com.CoInitialize.assert_not_called()

    def test_cancel_after_open_closes_without_saving(self):
        cancelled = threading.Event()
        self.ppt.Presentations.Open.side_effect = lambda *a, **k: (cancelled.set(), self.presentation)[1]
        with self.assertRaises(app.ConversionCancelled):
            app.text_to_image_ppt(str(self.source), cancel_event=cancelled)
        self.assertEqual(self.ppt.DisplayAlerts, 2)
        self.presentation.Close.assert_called_once()
        self.presentation.SaveAs.assert_not_called()
        self.com.CoUninitialize.assert_called_once()

    def test_invalid_saved_package_is_not_published(self):
        self.presentation.SaveAs.side_effect = lambda path, fmt: Path(path).write_bytes(b"broken")
        with self.assertRaises(app.UserVisibleError):
            self.convert()
        self.assertFalse(list(self.folder.glob("source_converted_*")))
        self.assertEqual(len(list(self.folder.glob(".ppt-pixel-pending-*"))), 1)

    def test_close_failure_keeps_recoverable_pending_file(self):
        self.presentation.Close.side_effect = RuntimeError("close failed")
        result = self.convert()
        self.assertIn(".ppt-pixel-pending-", result.output_path)
        report = json.loads(Path(result.report_path).read_text(encoding="utf-8"))
        self.assertEqual(result.state, "recovery_required")
        self.assertEqual(report["state"], result.state)
        self.assertFalse(list(self.folder.glob("source_converted_*")))

    def test_oversized_slide_stops_before_export(self):
        self.presentation.PageSetup.SlideWidth = 100_000
        with self.assertRaises(app.UserVisibleError) as raised:
            self.convert()
        self.assertEqual(raised.exception.kind, "render_limit")
        self.presentation.SaveAs.assert_not_called()
        self.assertEqual(self.ppt.DisplayAlerts, 2)
        self.presentation.Close.assert_called_once()

    def test_cancellation_between_shapes_does_not_save(self):
        cancelled = threading.Event()
        slide = MagicMock()
        slide.Shapes.Count = 2
        shape = MagicMock(HasTable=False, HasChart=False, HasSmartArt=False)
        shape.Type = 1
        slide.Shapes.side_effect = lambda index: shape
        self.presentation.Slides.Count = 1
        self.presentation.Slides.__iter__.return_value = iter([slide])
        with patch.object(app, "shape_to_cropped_picture", side_effect=lambda **kwargs: (cancelled.set(), True)[1]) as render:
            with self.assertRaises(app.ConversionCancelled):
                app.text_to_image_ppt(str(self.source), cancel_event=cancelled)
        render.assert_called_once()
        self.presentation.SaveAs.assert_not_called()
        self.assertEqual(self.ppt.DisplayAlerts, 2)
        self.presentation.Close.assert_called_once()
        self.assertEqual(self.source.read_bytes(), b"original")

    def test_report_failure_preserves_published_output_and_warns(self):
        with patch.object(app.os, "replace", side_effect=PermissionError("report denied")):
            result = self.convert()
        self.assertTrue(Path(result.output_path).is_file())
        self.assertEqual(result.report_path, "")
        self.assertEqual(result.state, "partial")
        self.assertEqual(result.issues[-1]["stage"], "report")
        self.assertFalse(list(self.folder.glob(".ppt-pixel-report-*")))

    def test_staging_permission_error_is_classified_and_closed(self):
        real_mkstemp = tempfile.mkstemp
        def create_temp(*args, **kwargs):
            if kwargs.get("prefix") == ".ppt-pixel-pending-":
                raise PermissionError("destination denied")
            return real_mkstemp(*args, **kwargs)
        with patch.object(app.tempfile, "mkstemp", side_effect=create_temp):
            with self.assertRaises(app.UserVisibleError) as raised:
                self.convert()
        self.assertEqual(raised.exception.kind, "file_save")
        self.assertEqual(self.ppt.DisplayAlerts, 2)
        self.presentation.Close.assert_called_once()

    def test_legacy_header_validation(self):
        legacy = self.folder / "legacy.ppt"
        legacy.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        app.validate_output(legacy, ".ppt")
        legacy.write_bytes(b"broken")
        with self.assertRaises(ValueError):
            app.validate_output(legacy, ".ppt")


if __name__ == "__main__":
    unittest.main()
