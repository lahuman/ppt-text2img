import logging
import subprocess
import sys
import tempfile
import unittest
from logging.handlers import RotatingFileHandler
from pathlib import Path
from unittest.mock import patch

from runtime_support import ConversionBusyError, configure_logging, conversion_slot


class RuntimeTests(unittest.TestCase):
    def test_thread_lock_released_after_failure(self):
        with self.assertRaisesRegex(RuntimeError, "injected"):
            with conversion_slot():
                with self.assertRaises(ConversionBusyError):
                    with conversion_slot():
                        self.fail("overlapping conversion")
                raise RuntimeError("injected")
        with conversion_slot():
            pass

    @unittest.skipUnless(sys.platform == "win32", "Windows interprocess lock")
    def test_another_process_cannot_acquire_slot(self):
        code = "from runtime_support import conversion_slot, ConversionBusyError\ntry:\n with conversion_slot(): pass\nexcept ConversionBusyError:\n raise SystemExit(23)"
        with conversion_slot():
            child = subprocess.run([sys.executable, "-B", "-c", code], capture_output=True, timeout=10)
            self.assertEqual(child.returncode, 23, child.stderr.decode(errors="replace"))
        child = subprocess.run([sys.executable, "-B", "-c", code], capture_output=True, timeout=10)
        self.assertEqual(child.returncode, 0, child.stderr.decode(errors="replace"))

    def test_logging_is_bounded_and_idempotent(self):
        root = logging.getLogger()
        original_level = root.level
        before = set(root.handlers)
        try:
            with tempfile.TemporaryDirectory() as directory:
                path = configure_logging(directory)
                configure_logging(directory)
                handlers = [h for h in root.handlers if h not in before]
                self.assertEqual(len(handlers), 1)
                self.assertIsInstance(handlers[0], RotatingFileHandler)
                self.assertEqual(handlers[0].backupCount, 3)
                self.assertEqual(handlers[0].maxBytes, 1024 * 1024)
                logging.info("diagnostic test")
                self.assertIn("diagnostic test", Path(path).read_text(encoding="utf-8"))
                for handler in handlers:
                    root.removeHandler(handler)
                    handler.close()
        finally:
            for handler in set(root.handlers) - before:
                root.removeHandler(handler)
                handler.close()
            root.setLevel(original_level)

    def test_logging_failure_does_not_block_startup(self):
        with patch("runtime_support.Path.mkdir", side_effect=PermissionError()):
            self.assertEqual(configure_logging("unwritable"), "")
