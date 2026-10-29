"""Opt-in integration checks; creates only temporary test presentations."""
import hashlib
import json
import os
import shutil
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from PIL import Image, ImageChops, ImageStat

import app


@unittest.skipUnless(os.environ.get("RUN_POWERPOINT_TESTS") == "1", "Requires installed PowerPoint")
class PowerPointTests(unittest.TestCase):
    def test_rendering_and_partial_failure(self):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        ppt = presentation = None
        temporary = tempfile.TemporaryDirectory()
        try:
            folder = temporary.name
            folder = Path(folder)
            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            ppt.DisplayAlerts = 1
            ppt.AutomationSecurity = 3
            presentation = ppt.Presentations.Add(False)
            presentation.PageSetup.SlideWidth = 720
            presentation.PageSetup.SlideHeight = 540
            for case in ("plain", "rotated", "overlap", "nested", "outside"):
                slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12)
                shape = slide.Shapes.AddTextbox(1, 60, 100, 300, 90)
                shape.TextFrame.TextRange.Text = "Font preservation 한글 테스트"
                shape.TextFrame.TextRange.Font.Name = "Malgun Gothic"
                shape.TextFrame.TextRange.Font.Size = 24
                if case == "rotated":
                    shape.Rotation = 37
                elif case == "overlap":
                    front = slide.Shapes.AddShape(1, 100, 110, 80, 60)
                    front.Fill.ForeColor.RGB = 255
                elif case == "nested":
                    companion = slide.Shapes.AddShape(1, 400, 100, 80, 60)
                    group = slide.Shapes.Range([shape.Name, companion.Name]).Group()
                    companion = slide.Shapes.AddShape(1, 500, 300, 40, 40)
                    slide.Shapes.Range([group.Name, companion.Name]).Group()
                elif case == "outside":
                    shape.Left = -45
                slide.Export(str(folder / (case + "_before.png")), "PNG", 1440, 1080)
            source = folder / "regression.pptx"
            presentation.SaveAs(str(source), 24)
            presentation.Close()
            presentation = ppt.Presentations.Add(False)
            presentation.Slides.Add(1, 12)
            original_hash = hashlib.sha256(source.read_bytes()).digest()
            # Match the GUI: conversion owns COM initialization on a worker thread.
            with ThreadPoolExecutor(max_workers=1) as worker:
                result = worker.submit(app.text_to_image_ppt, str(source)).result()
            self.assertEqual(result.converted_count, 5)
            self.assertEqual(result.issues, [])
            self.assertEqual(hashlib.sha256(source.read_bytes()).digest(), original_hash)
            self.assertEqual(presentation.Slides.Count, 1)
            presentation.Close()
            presentation = ppt.Presentations.Open(result.output_path, WithWindow=False)
            for index, case in enumerate(("plain", "rotated", "overlap", "nested", "outside"), 1):
                slide = presentation.Slides(index)
                path = folder / (case + "_after.png")
                slide.Export(str(path), "PNG", 1440, 1080)
                if os.environ.get("VERIFY_OUTPUT_DIR"):
                    evidence = Path(os.environ["VERIFY_OUTPUT_DIR"])
                    evidence.mkdir(parents=True, exist_ok=True)
                    for image_path in (folder / (case + "_before.png"), path):
                        shutil.copy2(image_path, evidence / image_path.name)
                with Image.open(folder / (case + "_before.png")) as before, Image.open(path) as after:
                    difference = ImageChops.difference(before.convert("RGB"), after.convert("RGB"))
                    error = sum(ImageStat.Stat(difference).mean) / 3
                    print(f"{case}: mean pixel difference {error:.4f}/255")
                    self.assertLess(error, 2.0, case)
                    white = Image.new("RGB", before.size, "white")
                    old_bounds = ImageChops.difference(before.convert("RGB"), white).getbbox()
                    new_bounds = ImageChops.difference(after.convert("RGB"), white).getbbox()
                    self.assertTrue(all(abs(a - b) <= 2 for a, b in zip(old_bounds, new_bounds)),
                                    (case, old_bounds, new_bounds))
                self.assertFalse(any(app.has_visible_text(slide.Shapes(i))
                                     for i in range(1, slide.Shapes.Count + 1)))
            presentation.Close()
            presentation = ppt.Presentations.Add(False)
            presentation.Slides.Add(1, 12)
            # The caller's PowerPoint session must remain usable after conversion.
            self.assertGreaterEqual(ppt.Presentations.Count, 0)
            from unittest.mock import patch
            with patch("app.shape_to_cropped_picture", side_effect=RuntimeError("injected export failure")):
                with ThreadPoolExecutor(max_workers=1) as worker:
                    partial = worker.submit(app.text_to_image_ppt, str(source)).result()
            self.assertEqual(partial.converted_count, 0)
            self.assertEqual(len(partial.issues), 5)
            report = json.loads(Path(partial.report_path).read_text(encoding="utf-8"))
            self.assertEqual(len(report["issues"]), 5)
            self.assertEqual(presentation.Slides.Count, 1)
            presentation.Close()
            presentation = ppt.Presentations.Open(str(source), WithWindow=False)
            legacy_source = folder / "legacy.ppt"
            presentation.SaveCopyAs(str(legacy_source), 1)
            with ThreadPoolExecutor(max_workers=1) as worker:
                with self.assertRaises(app.UserVisibleError):
                    worker.submit(app.text_to_image_ppt, str(source)).result()
            self.assertEqual(presentation.Slides.Count, 5)
            presentation.Close()
            presentation = ppt.Presentations.Add(False)
            slide = presentation.Slides.Add(1, 12)
            slide.Shapes.AddTable(2, 2, 50, 50, 200, 100)
            table_source = folder / "table.pptx"
            presentation.SaveAs(str(table_source), 24)
            presentation.Close()
            presentation = None
            with ThreadPoolExecutor(max_workers=1) as worker:
                unsupported = worker.submit(app.text_to_image_ppt, str(table_source)).result()
            self.assertEqual(unsupported.converted_count, 0)
            self.assertEqual(unsupported.issues[0]["stage"], "unsupported")
            with ThreadPoolExecutor(max_workers=1) as worker:
                legacy = worker.submit(app.text_to_image_ppt, str(legacy_source)).result()
            self.assertEqual(legacy.converted_count, 5)
            self.assertEqual(legacy.issues, [])
            self.assertEqual(Path(legacy.output_path).read_bytes()[:4], b"\xd0\xcf\x11\xe0")
        finally:
            if presentation is not None:
                presentation.Close()
            if ppt is not None:
                ppt.Quit()
            pythoncom.CoUninitialize()
            temporary.cleanup()


if __name__ == "__main__":
    unittest.main()
