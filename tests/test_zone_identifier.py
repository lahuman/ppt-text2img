import errno
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import conversion


@unittest.skipUnless(os.name == "nt", "Windows alternate data streams")
class ZoneIdentifierTests(unittest.TestCase):
    def test_invalid_argument_requires_volume_without_streams(self):
        with patch("conversion.open", side_effect=OSError(errno.EINVAL, "invalid argument")), \
             patch("conversion.win32file.GetVolumePathName", return_value="D:\\"), \
             patch("conversion.win32api.GetVolumeInformation") as information:
            information.return_value = ("", 0, 255, 0, "")
            conversion.copy_zone_identifier("source", "destination")
            information.return_value = ("", 0, 255, conversion.win32con.FILE_NAMED_STREAMS, "")
            with self.assertRaises(OSError):
                conversion.copy_zone_identifier("source", "destination")

    def test_zone_preserved_and_destination_error_propagated(self):
        with tempfile.TemporaryDirectory() as folder:
            source = str(Path(folder, "source.pptx"))
            destination = str(Path(folder, "copy.pptx"))
            Path(source).touch()
            Path(destination).touch()
            mark = b"[ZoneTransfer]\r\nZoneId=3\r\n"
            Path(source + ":Zone.Identifier").write_bytes(mark)
            conversion.copy_zone_identifier(source, destination)
            self.assertEqual(Path(destination + ":Zone.Identifier").read_bytes(), mark)
            with self.assertRaises(FileNotFoundError):
                conversion.copy_zone_identifier(source, str(Path(folder, "missing", "copy.pptx")))
