import os
import errno
import shutil
import tempfile
import logging
import json
import math
import time
import zipfile
from contextlib import contextmanager
from uuid import uuid4
from dataclasses import dataclass, field
from datetime import datetime

from embedded_fonts import inspect_fonts, write_replaced_copy
from PIL import Image
from runtime_support import ConversionBusyError, conversion_slot
from translations import LANG

try:
    import pythoncom
    import win32api
    import win32con
    import win32file
    import win32com.client
    COM_IMPORT_ERROR = None
except ImportError as exc:
    pythoncom = None
    win32com = None
    COM_IMPORT_ERROR = exc


MsoShapeTypeGroup = 6
msoFalse = 0
msoTrue = -1
ppShapeFormatPNG = 2
msoSendBackward = 3


class UserVisibleError(RuntimeError):
    def __init__(self, title, message, kind="generic"):
        super().__init__(message)
        self.title = title
        self.message = message
        self.kind = kind


class PowerPointRequiredError(UserVisibleError):
    def __init__(self, title, message):
        super().__init__(title, message, kind="powerpoint_required")


def make_powerpoint_required_error(texts):
    return PowerPointRequiredError(
        texts["dialog_powerpoint_title"],
        texts["error_powerpoint_required"]
    )


def looks_like_lock_error(message):
    lowered = message.lower()
    keywords = (
        "permission denied",
        "access denied",
        "sharing violation",
        "in use",
        "used by another process",
        "read-only",
        "locked",
        "사용 중",
        "권한",
        "잠금",
    )
    return any(keyword in lowered for keyword in keywords)


def looks_like_missing_file_error(message):
    lowered = message.lower()
    keywords = (
        "not found",
        "cannot find",
        "could not find",
        "파일을 찾을 수",
        "찾을 수 없습니다",
    )
    return any(keyword in lowered for keyword in keywords)


def classify_open_error(exc, texts):
    message = str(exc)

    if looks_like_missing_file_error(message):
        return UserVisibleError(
            texts["dialog_file_missing_title"],
            texts["error_file_not_found"],
            kind="file_open"
        )

    if looks_like_lock_error(message):
        return UserVisibleError(
            texts["dialog_file_locked_title"],
            texts["error_file_locked"],
            kind="file_open"
        )

    return UserVisibleError(
        texts["dialog_file_open_title"],
        texts["error_file_open_failed"],
        kind="file_open"
    )


def classify_save_error(exc, texts):
    message = str(exc)

    if looks_like_lock_error(message):
        return UserVisibleError(
            texts["dialog_file_save_locked_title"],
            texts["error_file_save_locked"],
            kind="file_save"
        )

    return UserVisibleError(
        texts["dialog_file_save_title"],
        texts["error_file_save_failed"],
        kind="file_save"
    )


@dataclass
class ConversionResult:
    output_path: str
    converted_count: int = 0
    issues: list = field(default_factory=list)
    report_path: str = ""
    metrics: dict = field(default_factory=dict)
    state: str = "complete"


class ConversionCancelled(RuntimeError):
    pass


class ConversionRecoveryError(RuntimeError):
    """Stop saving when temporary objects could not be removed safely."""


def log_failure(job_id, stage, exc, *, slide=None, shape=None):
    fields = [f"job={job_id}", f"stage={stage}"]
    if slide is not None:
        fields.extend((f"slide={slide}", f"shape={shape}"))
    fields.append(f"type={type(exc).__name__}")
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        fields.append(f"hresult=0x{hresult & 0xffffffff:08X}")
    info = getattr(exc, "excepinfo", None)
    if info:
        source, description, scode = info[1], info[2], info[5]
        if source:
            fields.append("source=" + json.dumps(source, ensure_ascii=False))
        if isinstance(scode, int):
            fields.append(f"scode=0x{scode & 0xffffffff:08X}")
        if description:
            fields.append("description=" + json.dumps(description, ensure_ascii=False))
    message = getattr(exc, "strerror", None) or str(exc)
    fields.append("message=" + json.dumps(message, ensure_ascii=False))
    logging.warning("%s", " ".join(fields))


def has_visible_text(shape):
    if not shape.HasTextFrame or not shape.TextFrame.HasText:
        return False
    text = shape.TextFrame.TextRange.Text
    return bool(text and text.strip())


def unsupported_shape(shape):
    return shape.HasTable or shape.HasChart or shape.HasSmartArt


def inspect_group(group, cancel_event):
    has_target = has_unsupported = False
    items = group.GroupItems
    for index in range(1, items.Count + 1):
        check_cancelled(cancel_event)
        child = items.Item(index)
        if child.Type == MsoShapeTypeGroup:
            target, unsupported = inspect_group(child, cancel_event)
            has_target |= target
            has_unsupported |= unsupported
        elif unsupported_shape(child):
            has_unsupported = True
        elif has_visible_text(child):
            has_target = True
    return has_target, has_unsupported


def ungroup_all_shapes(slide, issues, slide_index, cancel_event=None, job_id="-"):
    # Snapshot references: replacing a shape changes collection indices.
    pending = [(str(i), slide.Shapes(i)) for i in range(1, slide.Shapes.Count + 1)]
    while pending:
        check_cancelled(cancel_event)
        index, shape = pending.pop()
        try:
            if shape.Type == MsoShapeTypeGroup:
                has_target, has_unsupported = inspect_group(shape, cancel_event)
                if not has_target:
                    if has_unsupported:
                        issues.append({"slide": slide_index, "stage": "unsupported",
                                       "error": "Group contains table, chart or SmartArt"})
                    continue
                children = shape.Ungroup()
                pending.extend((f"{index}.{i}", children.Item(i)) for i in range(1, children.Count + 1))
        except ConversionCancelled:
            raise
        except Exception as exc:
            log_failure(job_id, "shape.ungroup", exc, slide=slide_index, shape=index)
            issues.append({"slide": slide_index, "stage": "ungroup", "error": str(exc)})


def crop_transparent_area(png_path, slide_width_pt, slide_height_pt):
    with Image.open(png_path, formats=("PNG",)) as source:
        img = source.convert("RGBA")
    with img:
        with img.getchannel("A") as alpha:
            bbox = alpha.getbbox()
        if not bbox:
            return None
        with img.crop(bbox) as cropped:
            cropped.save(png_path)
        scale_x = img.width / float(slide_width_pt)
        scale_y = img.height / float(slide_height_pt)
        return (bbox[0] / scale_x, bbox[1] / scale_y,
                (bbox[2] - bbox[0]) / scale_x, (bbox[3] - bbox[1]) / scale_y)


def ensure_runtime_requirements(texts):
    if os.name != "nt":
        raise UserVisibleError(
            texts["dialog_environment_title"],
            texts["error_windows_only"],
            kind="environment"
        )

    if pythoncom is None or win32com is None:
        logging.warning("PowerPoint runtime requirements missing: %s", COM_IMPORT_ERROR)
        raise make_powerpoint_required_error(texts)


def shape_to_cropped_picture(slide, shape, slide_width_pt, slide_height_pt, temp_png_path):
    original_z = shape.ZOrderPosition
    duplicate = anchor = group = new_shape = None
    try:
        operation = "shape.duplicate"
        duplicate = shape.Duplicate().Item(1)
        # Legacy PPT can give the duplicate the same name as the original.
        duplicate.Name = "__ppt_copy_" + uuid4().hex
        # PowerPoint offsets duplicates. Restore position without changing rotation.
        duplicate.Left = shape.Left
        duplicate.Top = shape.Top
        angle = math.radians(shape.Rotation)
        half_width = (abs(shape.Width * math.cos(angle)) + abs(shape.Height * math.sin(angle))) / 2
        half_height = (abs(shape.Width * math.sin(angle)) + abs(shape.Height * math.cos(angle))) / 2
        center_x, center_y = shape.Left + shape.Width / 2, shape.Top + shape.Height / 2
        # Keep rendered text overhang away from the export canvas edges.
        left = min(0, center_x - half_width) - 36
        top = min(0, center_y - half_height) - 36
        right = max(slide_width_pt, center_x + half_width) + 36
        bottom = max(slide_height_pt, center_y + half_height) + 36
        # Whole-point origins preserve the pixel grid at the 2x export scale.
        # Keep fractional canvases intact to preserve PowerPoint's rounding.
        if all(float(value).is_integer() for value in (left, top, right, bottom)):
            try:
                has_effects = (shape.Shadow.Visible or shape.ThreeD.Visible or shape.Glow.Radius
                               or shape.Reflection.Type or shape.SoftEdge.Radius)
                text = shape.TextFrame.TextRange
                text_left, text_top = text.BoundLeft, text.BoundTop
                text_right = text_left + text.BoundWidth
                text_bottom = text_top + text.BoundHeight
            except Exception:
                pass  # Some PowerPoint shapes do not expose rendered text bounds.

            else:
                if not has_effects:
                    left = max(left, math.floor(min(center_x - half_width, text_left)) - 36)
                    top = max(top, math.floor(min(center_y - half_height, text_top)) - 36)
                    right = min(right, math.ceil(max(center_x + half_width, text_right)) + 36)
                    bottom = min(bottom, math.ceil(max(center_y + half_height, text_bottom)) + 36)
        operation = "shape.anchor"
        anchor = slide.Shapes.AddShape(1, left, top, right - left, bottom - top)
        anchor.Name = "__ppt_anchor_" + uuid4().hex
        anchor.Line.Visible = msoFalse
        anchor.Fill.Visible = msoTrue
        anchor.Fill.Transparency = 1.0
        operation = "shape.group"
        group = slide.Shapes.Range([duplicate.Name, anchor.Name]).Group()
        # Use the actual group bounds, including objects outside the slide.
        origin_left, origin_top = group.Left, group.Top
        bounds_width, bounds_height = group.Width, group.Height
        # Render at a positive origin, then map back to the original coordinates.
        group.Left = 0
        group.Top = 0
        operation = "shape.export"
        group.Export(temp_png_path, ppShapeFormatPNG,
                     round(slide_width_pt * 2), round(slide_height_pt * 2))
        operation = "shape.crop"
        cropped = crop_transparent_area(temp_png_path, bounds_width, bounds_height)
        if not cropped:
            raise RuntimeError("No visible pixels found")
        left, top, width, height = cropped
        operation = "shape.insert"
        new_shape = slide.Shapes.AddPicture(
            temp_png_path, False, True, origin_left + left, origin_top + top, width, height
        )
        new_shape.Name = shape.Name + "_img"
        # Rotation is already rendered into the PNG; do not rotate it again.
        operation = "shape.delete_group"
        group.Delete()
        group = duplicate = anchor = None
        operation = "shape.stack"
        while new_shape.ZOrderPosition > original_z:
            previous_z = new_shape.ZOrderPosition
            new_shape.ZOrder(msoSendBackward)
            if new_shape.ZOrderPosition >= previous_z:
                raise RuntimeError("PowerPoint did not advance the stacking order")
        # Commit last: every operation above can fail without changing the original.
        operation = "shape.delete_original"
        shape.Delete()
        return True
    except Exception as primary:
        primary.shape_operation = operation
        cleanup_errors = []
        objects = [new_shape, group] if group is not None else [new_shape, duplicate, anchor]
        for obj in objects:
            if obj is not None:
                try:
                    obj.Delete()
                except Exception as exc:
                    cleanup_errors.append(exc)
        if cleanup_errors:
            recovery = ConversionRecoveryError("Temporary shape cleanup failed; output was not saved")
            recovery.cleanup_errors = cleanup_errors
            raise recovery from primary
        raise

def check_cancelled(cancel_event):
    if cancel_event is not None and cancel_event.is_set():
        raise ConversionCancelled()


def text_to_image_ppt(input_ppt, progress_callback=None, texts=None, cancel_event=None, saving_callback=None, font_choice=None):
    texts = texts or LANG["ko"]
    check_cancelled(cancel_event)
    try:
        with conversion_slot():
            return _convert_presentation(input_ppt, progress_callback, texts, cancel_event, saving_callback, font_choice)
    except ConversionBusyError as exc:
        raise UserVisibleError(texts["dialog_error_title"], texts["error_busy"], kind="busy") from exc


@contextmanager
def suppress_alerts(powerpoint, cleanup_issues, job_id, phase):
    previous = powerpoint.DisplayAlerts
    powerpoint.DisplayAlerts = 1
    try:
        yield
    finally:
        try:
            powerpoint.DisplayAlerts = previous
        except Exception as exc:
            log_failure(job_id, phase + ".restore_alerts", exc)
            cleanup_issues.append({"stage": "restore_alerts", "error": str(exc)})


def copy_zone_identifier(source, destination):
    try:
        zone = open(source + ":Zone.Identifier", "rb")
    except FileNotFoundError:
        return
    except OSError as exc:
        if getattr(exc, "winerror", None) in (1, 50):  # Unsupported stream operation
            return
        if exc.errno == errno.EINVAL:
            volume = win32file.GetVolumePathName(os.path.realpath(source))
            flags = win32api.GetVolumeInformation(volume)[3]
            if not flags & win32con.FILE_NAMED_STREAMS:
                return
        raise
    with zone, open(destination + ":Zone.Identifier", "wb") as target:
        shutil.copyfileobj(zone, target)


def _convert_presentation(input_ppt, progress_callback=None, texts=None, cancel_event=None, saving_callback=None, font_choice=None):
    texts = texts or LANG["ko"]
    base_name, ext = os.path.splitext(input_ppt)
    if ext.lower() not in (".ppt", ".pptx"):
        raise UserVisibleError(
            texts["dialog_invalid_file_title"],
            texts["error_invalid_file_type"],
            kind="invalid_file"
        )

    ensure_runtime_requirements(texts)

    abs_input_path = os.path.abspath(input_ppt)
    if not os.path.exists(abs_input_path):
        raise UserVisibleError(
            texts["dialog_file_missing_title"],
            texts["error_file_not_found"],
            kind="file_open"
        )

    pythoncom.CoInitialize()

    powerpoint = None
    presentation = None
    temp_png_path = None
    owns_powerpoint = False
    original_security = None
    cleanup_issues = []
    result = None
    font_copy = None
    staged_path = None
    saved = False

    started = time.perf_counter()
    job_id = uuid4().hex
    logging.info("Job %s started", job_id)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    output_ppt = "%s_converted_%s_%s%s" % (base_name, timestamp, job_id[:8], ext)
    abs_output_path = os.path.abspath(output_ppt)

    try:
        open_path = abs_input_path
        if font_choice and ext.lower() == ".pptx" and zipfile.is_zipfile(abs_input_path):
            with zipfile.ZipFile(abs_input_path) as package:
                check = lambda: check_cancelled(cancel_event)
                restricted = inspect_fonts(package, check)
                if restricted:
                    replacement = font_choice(restricted)
                    check()
                    if not replacement:
                        raise ConversionCancelled()
                    with tempfile.NamedTemporaryFile(dir=os.path.dirname(abs_input_path),
                                                     prefix=".ppt-pixel-font-", suffix=ext,
                                                     delete=False) as working:
                        font_copy = working.name
                        write_replaced_copy(package, working, restricted, replacement, check)
                    if os.name == "nt":
                        copy_zone_identifier(abs_input_path, font_copy)
                    open_path = font_copy
        temp_fd, temp_png_path = tempfile.mkstemp(prefix="ppt_text_", suffix=".png")
        os.close(temp_fd)

        try:
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
                owns_powerpoint = True
        except Exception as exc:
            logging.warning("PowerPoint launch failed (%s)", type(exc).__name__)
            raise make_powerpoint_required_error(texts)
        original_security = powerpoint.AutomationSecurity
        for opened in powerpoint.Presentations:
            if os.path.normcase(os.path.abspath(opened.FullName)) == os.path.normcase(abs_input_path):
                raise UserVisibleError(texts["dialog_file_locked_title"], texts["error_file_locked"],
                                       kind="file_open")
        powerpoint.AutomationSecurity = 3  # Disable macros in programmatically opened files.

        try:
            with suppress_alerts(powerpoint, cleanup_issues, job_id, "open"):
                presentation = powerpoint.Presentations.Open(open_path, ReadOnly=True, Untitled=True, WithWindow=False)
        except Exception as exc:
            logging.warning("Presentation open failed (%s)", type(exc).__name__)
            raise classify_open_error(exc, texts)
        finally:
            try:
                powerpoint.AutomationSecurity = original_security
            except Exception as exc:
                cleanup_issues.append({"stage": "restore_security", "error": str(exc)})

        check_cancelled(cancel_event)
        total_slides = presentation.Slides.Count
        slide_width_pt = presentation.PageSetup.SlideWidth
        slide_height_pt = presentation.PageSetup.SlideHeight
        dimensions = (float(slide_width_pt), float(slide_height_pt))
        if (not all(math.isfinite(v) and v > 0 for v in dimensions)
                or min(round(v * 2) for v in dimensions) < 1
                or round(dimensions[0] * 2) * round(dimensions[1] * 2) > 25_000_000):
            raise UserVisibleError(texts["dialog_error_title"], texts["error_render_limit"], kind="render_limit")

        if progress_callback:
            progress_callback(0, total_slides, texts["status_start"])

        result = ConversionResult(abs_output_path)
        render_seconds = 0.0

        for slide_index, slide in enumerate(presentation.Slides, start=1):
            check_cancelled(cancel_event)
            if progress_callback:
                progress_callback(
                    slide_index - 1,
                    total_slides,
                    texts["status_slide_processing"].format(
                        current=slide_index,
                        total=total_slides
                    )
                )

            # 그룹 해제 실패는 확인 필요 항목으로 기록한다.
            ungroup_all_shapes(slide, result.issues, slide_index, cancel_event, job_id)

            shapes = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]
            next_update = 0.0
            for i, shape in enumerate(shapes, start=1):
                check_cancelled(cancel_event)
                now = time.monotonic()
                if progress_callback and (now >= next_update or i == len(shapes)):
                    progress_callback(slide_index - 1 + (i - 1) / len(shapes), total_slides,
                                      texts["shape_progress"].format(current=slide_index, total=total_slides,
                                                                     shape=i, count=len(shapes)))
                    next_update = now + 0.35
                stage = "shape.inspect"
                try:
                    if shape.Type == MsoShapeTypeGroup:
                        continue
                    if unsupported_shape(shape):
                        result.issues.append({"slide": slide_index, "shape_index": i,
                                              "stage": "unsupported",
                                              "error": "Table, chart or SmartArt requires manual review"})
                        continue

                    if not has_visible_text(shape):
                        continue

                    stage = "shape.render"
                    render_started = time.perf_counter()
                    try:
                        ok = shape_to_cropped_picture(
                            slide=slide,
                            shape=shape,
                            slide_width_pt=slide_width_pt,
                            slide_height_pt=slide_height_pt,
                            temp_png_path=temp_png_path
                        )
                    finally:
                        render_seconds += time.perf_counter() - render_started
                    if ok:
                        result.converted_count += 1

                except ConversionRecoveryError as exc:
                    primary = exc.__cause__ or exc
                    log_failure(job_id, getattr(primary, "shape_operation", stage), primary, slide=slide_index, shape=i)
                    for error in getattr(exc, "cleanup_errors", ()):
                        log_failure(job_id, "shape.rollback", error, slide=slide_index, shape=i)
                    raise
                except Exception as e:
                    result.issues.append({"slide": slide_index, "shape_index": i,
                                          "stage": "convert", "error": str(e)})
                    log_failure(job_id, getattr(e, "shape_operation", stage), e, slide=slide_index, shape=i)

            if progress_callback:
                progress_callback(
                    slide_index,
                    total_slides,
                    texts["status_slide_done"].format(
                        current=slide_index,
                        total=total_slides
                    )
                )

        check_cancelled(cancel_event)
        if saving_callback:
            saving_callback()
        if progress_callback:
            progress_callback(total_slides, total_slides, texts["status_saving"])
        save_started = time.perf_counter()
        stage = "save.create_temp"
        try:
            staged_fd, staged_path = tempfile.mkstemp(prefix=".ppt-pixel-pending-", suffix=ext,
                                                      dir=os.path.dirname(abs_output_path))
            os.close(staged_fd)
            stage = "save.configure_alerts"
            with suppress_alerts(powerpoint, cleanup_issues, job_id, "save"):
                stage = "save.SaveAs"
                presentation.SaveAs(staged_path, 1 if ext.lower() == ".ppt" else 24)
            saved = True
            result.metrics.update(render_seconds=round(render_seconds, 4),
                                  save_seconds=round(time.perf_counter() - save_started, 4))
        except Exception as exc:
            log_failure(job_id, stage, exc)
            raise classify_save_error(exc, texts)


    finally:
        def cleanup(stage, action):
            try:
                action()
            except Exception as exc:
                cleanup_issues.append({"stage": stage, "error": str(exc)})
                log_failure(job_id, "cleanup." + stage, exc)

        if presentation is not None:
            cleanup("close", presentation.Close)
        if powerpoint is not None:
            if original_security is not None:
                cleanup("restore_security", lambda: setattr(powerpoint, "AutomationSecurity", original_security))
            if owns_powerpoint:
                def quit_owned_session():
                    if powerpoint.Presentations.Count == 0:
                        powerpoint.Quit()
                cleanup("quit", quit_owned_session)
        if temp_png_path:
            cleanup("temp_cleanup", lambda: os.remove(temp_png_path))
        if font_copy:
            cleanup("font_copy_cleanup", lambda: os.remove(font_copy))
        if staged_path and not saved:
            cleanup("staged_cleanup", lambda: os.remove(staged_path))
        cleanup("com_uninitialize", pythoncom.CoUninitialize)

    result.issues.extend(cleanup_issues)
    publish_output(result, staged_path, texts, job_id)
    result.metrics["elapsed_seconds"] = round(time.perf_counter() - started, 4)
    write_report(result, job_id)
    if result.issues and result.state == "complete":
        result.state = "partial"
    logging.info("Job %s finished state=%s converted=%s issues=%s elapsed=%.3fs",
                 job_id, result.state, result.converted_count, len(result.issues), result.metrics["elapsed_seconds"])
    if progress_callback:
        message = (texts["status_done_count"].format(count=result.converted_count)
                   if result.state == "complete" else texts["status_" + result.state])
        progress_callback(total_slides, total_slides, message)
    return result


def publish_output(result, staged_path, texts, job_id):
    if any(issue["stage"] == "close" for issue in result.issues):
        # Keep a recoverable file; never publish a file still held by PowerPoint.
        result.output_path = staged_path
        result.state = "recovery_required"
    else:
        stage = "save.validate"
        try:
            validate_output(staged_path, os.path.splitext(result.output_path)[1])
            stage = "save.rename"
            os.rename(staged_path, result.output_path)
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            log_failure(job_id, stage, exc)
            raise UserVisibleError(texts["dialog_file_save_title"],
                                   texts["error_file_save_failed"] + "\n" + staged_path,
                                   kind="file_save") from exc
        result.state = "partial" if result.issues else "complete"


def write_report(result, job_id):
    result.report_path = result.output_path + ".report.json"
    report_temp = None
    try:
        report_fd, report_temp = tempfile.mkstemp(prefix=".ppt-pixel-report-", suffix=".json",
                                                dir=os.path.dirname(result.output_path))
        with os.fdopen(report_fd, "w", encoding="utf-8") as report:
            json.dump({"schema_version": 1, "job_id": job_id, "state": result.state,
                       "metrics": result.metrics, "converted_count": result.converted_count,
                       "issues": result.issues}, report, ensure_ascii=False, indent=2)
        os.replace(report_temp, result.report_path)
    except OSError as exc:
        result.report_path = ""
        result.issues.append({"stage": "report", "error": str(exc)})
        logging.warning("Job %s report failed (%s)", job_id, type(exc).__name__)
    finally:
        if report_temp and os.path.exists(report_temp):
            try:
                os.remove(report_temp)
            except OSError as exc:
                result.issues.append({"stage": "report_cleanup", "error": str(exc)})


def validate_output(path, extension):
    if extension.lower() == ".pptx":
        with zipfile.ZipFile(path) as archive:
            if not {"[Content_Types].xml", "ppt/presentation.xml"}.issubset(archive.namelist()):
                raise ValueError("Missing PowerPoint package parts")
            if archive.testzip() is not None:
                raise ValueError("Corrupt PowerPoint package")
    else:
        with open(path, "rb") as stream:
            if stream.read(8) != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                raise ValueError("Invalid legacy PowerPoint header")
