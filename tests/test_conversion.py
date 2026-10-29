import tempfile
import queue
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from PIL import Image

import app


class ConversionTests(unittest.TestCase):
    def test_transparent_crop_maps_pixels_to_points(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "crop.png"
            image = Image.new("RGBA", (200, 100))
            image.paste((0, 0, 0, 255), (20, 10, 80, 50))
            image.save(path)
            self.assertEqual(app.crop_transparent_area(path, 100, 50), (10, 5, 30, 20))
            with Image.open(path) as cropped:
                self.assertEqual(cropped.size, (60, 40))

    def test_empty_image_is_not_a_conversion(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "empty.png"
            Image.new("RGBA", (100, 100)).save(path)
            self.assertIsNone(app.crop_transparent_area(path, 100, 100))

    def make_shapes(self):
        original = Mock(Name="original", Left=20, Top=30, Width=100, Height=40,
                        Rotation=0, ZOrderPosition=1)
        duplicate = Mock(Name="copy")
        anchor = Mock(Name="anchor")
        group = Mock(Left=0, Top=0, Width=720, Height=540)
        picture = Mock(ZOrderPosition=1)
        original.Duplicate.return_value.Item.return_value = duplicate
        slide = Mock()
        slide.Shapes.AddShape.return_value = anchor
        slide.Shapes.Range.return_value.Group.return_value = group
        slide.Shapes.AddPicture.return_value = picture
        return original, slide, duplicate, anchor, group, picture

    @patch("app.crop_transparent_area", return_value=(20, 30, 100, 40))
    def test_export_failure_preserves_original(self, crop):
        original, slide, duplicate, anchor, group, picture = self.make_shapes()
        group.Export.side_effect = RuntimeError("export failed")
        with self.assertRaisesRegex(RuntimeError, "export failed"):
            app.shape_to_cropped_picture(slide, original, 720, 540, "unused.png")
        original.Delete.assert_not_called()
        group.Delete.assert_called_once()
        slide.Shapes.AddPicture.assert_not_called()

    @patch("app.crop_transparent_area", return_value=(20, 30, 100, 40))
    def test_original_delete_failure_removes_replacement(self, crop):
        original, slide, duplicate, anchor, group, picture = self.make_shapes()
        original.Delete.side_effect = RuntimeError("delete failed")
        with self.assertRaisesRegex(RuntimeError, "delete failed"):
            app.shape_to_cropped_picture(slide, original, 720, 540, "unused.png")
        picture.Delete.assert_called_once()
        self.assertEqual(group.Delete.call_count, 1)

    @patch("app.crop_transparent_area", return_value=(20, 30, 100, 40))
    def test_cleanup_failure_stops_document_save(self, crop):
        original, slide, duplicate, anchor, group, picture = self.make_shapes()
        group.Export.side_effect = RuntimeError("export failed")
        group.Delete.side_effect = RuntimeError("cleanup failed")
        with self.assertRaises(app.ConversionRecoveryError):
            app.shape_to_cropped_picture(slide, original, 720, 540, "unused.png")
        original.Delete.assert_not_called()

    def test_text_inspection_error_is_not_hidden(self):
        shape = Mock()
        type(shape).HasTextFrame = property(lambda self: (_ for _ in ()).throw(RuntimeError("COM")))
        with self.assertRaises(RuntimeError):
            app.has_visible_text(shape)

    @patch("app.messagebox.showwarning")
    def test_close_during_conversion_keeps_window_open(self, warning):
        ui = app.App.__new__(app.App)
        ui.root = Mock()
        ui.lang_code = "ko"
        ui.is_converting = True
        ui.on_close()
        ui.root.destroy.assert_not_called()
        warning.assert_called_once()

    @patch("app.messagebox.showinfo")
    @patch("app.messagebox.showwarning")
    def test_partial_conversion_shows_warning_instead_of_success(self, warning, info):
        ui = app.App.__new__(app.App)
        ui.lang_code = "ko"
        ui.is_converting = False
        ui.progress_bar = {"maximum": 5}
        ui.lbl_percent = Mock()
        ui.ui_event_queue = queue.Queue()
        ui.ui_event_queue.put({"type": "done", "converted_count": 4,
                               "issues": [{"slide": 2}], "output_path": "out.pptx",
                               "report_path": "out.pptx.report.json"})
        ui.set_status = Mock()
        ui.set_busy = Mock()
        ui.process_ui_events()
        warning.assert_called_once()
        info.assert_not_called()

    @patch("app.filedialog.askopenfilename", return_value="C:/slides/talk.pptx")
    @patch("app.messagebox.askokcancel")
    def test_file_selection_waits_for_explicit_conversion(self, confirm, dialog):
        ui = app.App.__new__(app.App)
        ui.lang_code = "ko"
        ui.is_converting = False
        ui.update_selected_file_label = Mock()
        ui.update_progress = Mock()
        ui.select_file()
        self.assertEqual(ui.selected_file, "C:/slides/talk.pptx")
        confirm.assert_not_called()

    @patch("app.messagebox.askokcancel", return_value=False)
    def test_cancel_conversion_does_not_start_worker(self, confirm):
        ui = app.App.__new__(app.App)
        ui.lang_code = "ko"
        ui.is_converting = False
        ui.selected_file = "C:/slides/talk.pptx"
        ui.set_status = Mock()
        with patch("app.threading.Thread") as worker:
            ui.start_conversion()
        worker.assert_not_called()


if __name__ == "__main__":
    unittest.main()
