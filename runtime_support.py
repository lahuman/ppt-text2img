import logging
import os
import tempfile
import threading
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler
from pathlib import Path


class ConversionBusyError(RuntimeError):
    pass


_conversion_lock = threading.Lock()


@contextmanager
def conversion_slot():
    """Serialize threads and Windows processes sharing the user's temp directory."""
    if not _conversion_lock.acquire(blocking=False):
        raise ConversionBusyError("Another conversion is running")
    handle = None
    locked = False
    try:
        if os.name == "nt":
            import msvcrt
            handle = open(Path(tempfile.gettempdir()) / "ppt-pixel-conversion.lock", "a+b")
            handle.seek(0, os.SEEK_END)
            if handle.tell() == 0:
                handle.write(b"0")
                handle.flush()
            handle.seek(0)
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                raise ConversionBusyError("Another conversion is running") from exc
            locked = True
        yield
    finally:
        try:
            if locked:
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        finally:
            try:
                if handle is not None:
                    handle.close()
            finally:
                _conversion_lock.release()


def configure_logging(directory=None):
    """Logging is optional: an unwritable log directory must not block startup."""
    try:
        directory = Path(directory or Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "PPT-Pixel" / "logs")
        directory.mkdir(parents=True, exist_ok=True)
        path = (directory / "app.log").resolve()
        root = logging.getLogger()
        if not any(isinstance(h, RotatingFileHandler) and h.baseFilename == str(path) for h in root.handlers):
            handler = RotatingFileHandler(path, maxBytes=1024 * 1024, backupCount=3, encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
            root.addHandler(handler)
        root.setLevel(logging.INFO)
        return str(path)
    except OSError:
        return ""
