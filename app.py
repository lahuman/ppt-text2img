import os
import queue
import tempfile
import logging
import threading
import json
import math
from uuid import uuid4
from dataclasses import dataclass, field
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image

APP_VERSION = "1.2"

try:
    import pythoncom
    import win32com.client
    COM_IMPORT_ERROR = None
except ImportError as exc:
    pythoncom = None
    win32com = None
    COM_IMPORT_ERROR = exc

LANG = {
    "ko": {
        "app_title": "PPT 글씨 -> 이미지 변환기",
        "main_title": "글꼴은 달라도,\n디자인은 그대로.",
        "subtitle": "텍스트를 이미지로 바꿔, 다른 PC에서도 의도한 모습으로 전달하세요.",
        "lang_label": "언어",
        "btn_select": "파일 찾아보기",
        "btn_convert": "이미지로 변환하기  →",
        "file_heading": "변환할 프레젠테이션을 선택하세요",
        "file_formats": "PowerPoint 프레젠테이션 · .pptx / .ppt",
        "file_change": "다른 파일 선택",
        "workspace": "프레젠테이션 변환",
        "side_title": "전달하고 싶은 모습,\n그대로 남기세요.",
        "side_body": "글꼴이 없는 PC에서도\n발표의 디테일을 지킬 수 있도록.",
        "step_select": "프레젠테이션 선택",
        "step_convert": "텍스트를 이미지로 변환",
        "step_save": "새 파일로 안전하게 저장",
        "local_note": "파일은 내 PC에서 처리됩니다",
        "ready_hint": "파일을 선택하면 변환을 시작할 수 있습니다.",
        "selected_hint": "준비되었습니다. 아래 버튼으로 변환을 시작하세요.",
        "compact_warning": "변환할 파일을 PowerPoint에서 닫아 주세요.\n변환 후에는 텍스트를 직접 편집할 수 없습니다.",
        "requirements": "Windows · Microsoft PowerPoint 필요",
        "btn_busy": "변환 중...",
        "file_label": "선택한 파일",
        "file_placeholder": "아직 파일을 선택하지 않았습니다.",
        "output_hint": "결과 파일은 원본과 같은 폴더에 새 파일로 저장됩니다.",
        "warning_title": "변환 전 확인 사항",
        "warning_body": (
            "1. Microsoft PowerPoint를 완전히 종료하세요.\n"
            "2. 변환할 PPT 파일도 PowerPoint에서 닫혀 있어야 합니다.\n"
            "3. 원본 보존을 위해 복사본으로 작업하는 것을 권장합니다.\n"
            "4. 변환이 끝날 때까지 같은 파일을 다시 열지 마세요."
        ),
        "progress_title": "진행 상황",
        "status_ready": "대기 중",
        "status_preparing": "프로그램을 준비 중입니다...",
        "status_start": "변환을 시작합니다...",
        "status_slide_processing": "슬라이드 {current}/{total} 처리 중...",
        "status_slide_done": "슬라이드 {current}/{total} 완료",
        "status_done": "변환이 모두 완료되었습니다!",
        "status_partial": "일부 항목을 변환하지 못했습니다. 결과 보고서를 확인하세요.",
        "dialog_partial_title": "일부 변환 완료",
        "dialog_partial_msg": "변환: {count}개 / 확인 필요: {failed}건\n결과 파일:\n{path}\n\n보고서:\n{report}",
        "dialog_close_busy": "변환 중입니다. 작업이 끝난 후 창을 닫아 주세요.",
        "status_done_count": "완료: 텍스트 {count}개 변환",
        "status_error": "오류 발생!",
        "dialog_select": "변환할 PPT 파일을 선택하세요",
        "dialog_done_title": "완료",
        "dialog_done_msg": "변환이 완료되었습니다.\n변환된 텍스트 수: {count}\n저장 위치:\n{path}",
        "dialog_confirm_title": "변환 시작 확인",
        "dialog_confirm_msg": (
            "선택한 파일: {name}\n\n"
            "변환 전 확인 사항:\n"
            "1. Microsoft PowerPoint를 종료했습니다.\n"
            "2. 선택한 PPT 파일이 열려 있지 않습니다.\n"
            "3. 변환 중에는 같은 파일을 다시 열지 않겠습니다.\n\n"
            "계속하시겠습니까?"
        ),
        "dialog_error_title": "처리 중 문제가 발생했습니다",
        "dialog_error_msg": (
            "처리 중 예상하지 못한 문제가 발생했습니다.\n"
            "잠시 후 다시 시도해 주세요.\n\n"
            "상세 내용:\n{error}"
        ),
        "dialog_environment_title": "실행 환경 확인 필요",
        "dialog_powerpoint_title": "Microsoft PowerPoint 필요",
        "dialog_invalid_file_title": "파일 형식 확인 필요",
        "dialog_file_missing_title": "파일을 찾을 수 없습니다",
        "dialog_file_open_title": "파일을 열 수 없습니다",
        "dialog_file_locked_title": "파일이 사용 중입니다",
        "dialog_file_save_title": "결과 파일을 저장할 수 없습니다",
        "dialog_file_save_locked_title": "결과 파일이 사용 중입니다",
        "error_windows_only": "이 프로그램은 Windows에서만 실행할 수 있습니다.\nWindows PC에서 다시 실행해 주세요.",
        "error_powerpoint_required": (
            "이 기능을 사용하려면 Microsoft PowerPoint 정식 버전이 필요합니다.\n"
            "PowerPoint가 설치되어 있는지 확인한 후 다시 시도해 주세요."
        ),
        "error_invalid_file_type": (
            "PowerPoint 파일(.ppt 또는 .pptx)만 변환할 수 있습니다.\n"
            "다른 형식의 파일은 지원되지 않습니다."
        ),
        "error_file_not_found": (
            "선택한 파일을 찾을 수 없습니다.\n"
            "파일이 이동되었거나 삭제되었는지 확인한 후 다시 선택해 주세요."
        ),
        "error_file_open_failed": (
            "선택한 파일을 열 수 없습니다.\n"
            "파일이 손상되었거나 PowerPoint에서 지원하지 않는 형식일 수 있습니다.\n"
            "다른 PowerPoint 파일로 다시 시도해 주세요."
        ),
        "error_file_locked": (
            "선택한 파일이 다른 프로그램에서 사용 중입니다.\n"
            "Microsoft PowerPoint와 관련 프로그램을 모두 닫은 후 다시 시도해 주세요."
        ),
        "error_file_save_failed": (
            "변환된 파일을 저장하지 못했습니다.\n"
            "저장 폴더 권한이나 디스크 상태를 확인한 후 다시 시도해 주세요."
        ),
        "error_file_save_locked": (
            "같은 이름의 결과 파일이 열려 있거나 폴더에 쓰기 권한이 없습니다.\n"
            "열려 있는 파일을 닫고 다시 시도해 주세요."
        ),
        "lang_ko": "한국어",
        "lang_en": "English",
    },
    "en": {
        "app_title": "PPT Text to Image Converter",
        "main_title": "Different fonts.\nThe same impression.",
        "subtitle": "Turn text into images so your presentation travels with its design intact.",
        "lang_label": "Language",
        "btn_select": "Browse files",
        "btn_convert": "Convert to images  →",
        "file_heading": "Choose your presentation",
        "file_formats": "PowerPoint presentation · .pptx / .ppt",
        "file_change": "Choose another file",
        "workspace": "PRESENTATION CONVERTER",
        "side_title": "Keep the details.\nShare your vision.",
        "side_body": "Make your presentation look right,\neven on a PC without your fonts.",
        "step_select": "Choose a presentation",
        "step_convert": "Turn text into images",
        "step_save": "Save as a separate file",
        "local_note": "Your files stay on your PC",
        "ready_hint": "Choose a file to get started.",
        "selected_hint": "Ready when you are. Start the conversion below.",
        "compact_warning": "Close the source file in PowerPoint before converting.\nConverted text will no longer be directly editable.",
        "requirements": "Requires Windows and Microsoft PowerPoint",
        "btn_busy": "Converting...",
        "file_label": "Selected file",
        "file_placeholder": "No file selected yet.",
        "output_hint": "The converted file will be saved next to the original file.",
        "warning_title": "Before You Convert",
        "warning_body": (
            "1. Close Microsoft PowerPoint completely.\n"
            "2. Make sure the PPT file is not open in PowerPoint.\n"
            "3. Using a copy of the original file is recommended.\n"
            "4. Do not reopen the same file until the conversion finishes."
        ),
        "progress_title": "Progress",
        "status_ready": "Ready",
        "status_preparing": "Preparing the program...",
        "status_start": "Starting conversion...",
        "status_slide_processing": "Processing slide {current}/{total}...",
        "status_slide_done": "Slide {current}/{total} completed",
        "status_done": "Conversion completed!",
        "status_partial": "Some items could not be converted. Check the report.",
        "dialog_partial_title": "Partially Converted",
        "dialog_partial_msg": "Converted: {count} / Needs review: {failed}\nOutput:\n{path}\n\nReport:\n{report}",
        "dialog_close_busy": "Conversion is running. Please close the window after it finishes.",
        "status_done_count": "Done: converted {count} text items",
        "status_error": "Error occurred!",
        "dialog_select": "Select the PPT file to convert",
        "dialog_done_title": "Done",
        "dialog_done_msg": "Conversion completed.\nConverted text count: {count}\nSaved to:\n{path}",
        "dialog_confirm_title": "Confirm Conversion",
        "dialog_confirm_msg": (
            "Selected file: {name}\n\n"
            "Before converting:\n"
            "1. Microsoft PowerPoint has been closed.\n"
            "2. The selected PPT file is not open.\n"
            "3. You will not reopen the file during conversion.\n\n"
            "Do you want to continue?"
        ),
        "dialog_error_title": "Something Went Wrong",
        "dialog_error_msg": (
            "An unexpected problem occurred while processing the file.\n"
            "Please try again in a moment.\n\n"
            "Details:\n{error}"
        ),
        "dialog_environment_title": "Unsupported Environment",
        "dialog_powerpoint_title": "Microsoft PowerPoint Required",
        "dialog_invalid_file_title": "Unsupported File Type",
        "dialog_file_missing_title": "File Not Found",
        "dialog_file_open_title": "Unable to Open File",
        "dialog_file_locked_title": "File Is In Use",
        "dialog_file_save_title": "Unable to Save the Converted File",
        "dialog_file_save_locked_title": "Converted File Is In Use",
        "error_windows_only": "This program can only run on Windows.\nPlease run it on a Windows PC.",
        "error_powerpoint_required": (
            "A full Microsoft PowerPoint installation is required to use this feature.\n"
            "Please make sure PowerPoint is installed and try again."
        ),
        "error_invalid_file_type": (
            "Only PowerPoint files (.ppt or .pptx) can be converted.\n"
            "Other file types are not supported."
        ),
        "error_file_not_found": (
            "The selected file could not be found.\n"
            "Please check whether it was moved or deleted, then select it again."
        ),
        "error_file_open_failed": (
            "The selected file could not be opened.\n"
            "It may be damaged or not supported by PowerPoint.\n"
            "Please try another PowerPoint file."
        ),
        "error_file_locked": (
            "The selected file is currently being used by another program.\n"
            "Please close Microsoft PowerPoint and any related apps, then try again."
        ),
        "error_file_save_failed": (
            "The converted file could not be saved.\n"
            "Please check the destination folder permissions or disk status and try again."
        ),
        "error_file_save_locked": (
            "A file with the same name may already be open, or the folder may not allow writing.\n"
            "Please close the file and try again."
        ),
        "lang_ko": "한국어",
        "lang_en": "English",
    }
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

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


class ConversionRecoveryError(RuntimeError):
    """Stop saving when temporary objects could not be removed safely."""


def has_visible_text(shape):
    if not shape.HasTextFrame or not shape.TextFrame.HasText:
        return False
    text = shape.TextFrame.TextRange.Text
    return bool(text and text.strip())


def ungroup_all_shapes(slide, issues, slide_index):
    # Snapshot references: replacing a shape changes collection indices.
    pending = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]
    while pending:
        shape = pending.pop()
        try:
            if shape.Type == MsoShapeTypeGroup:
                children = shape.Ungroup()
                pending.extend(children.Item(i) for i in range(1, children.Count + 1))
        except Exception as exc:
            issues.append({"slide": slide_index, "stage": "ungroup", "error": str(exc)})


def crop_transparent_area(png_path, slide_width_pt, slide_height_pt):
    with Image.open(png_path).convert("RGBA") as img:
        alpha = img.getchannel("A")
        bbox = alpha.getbbox()

        if not bbox:
            return None

        cropped = img.crop(bbox)
        cropped.save(png_path)

        scale_x = img.width / float(slide_width_pt)
        scale_y = img.height / float(slide_height_pt)

        left_pt = bbox[0] / scale_x
        top_pt = bbox[1] / scale_y
        width_pt = (bbox[2] - bbox[0]) / scale_x
        height_pt = (bbox[3] - bbox[1]) / scale_y

        return left_pt, top_pt, width_pt, height_pt


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
        anchor = slide.Shapes.AddShape(1, left, top, right - left, bottom - top)
        anchor.Name = "__ppt_anchor_" + uuid4().hex
        anchor.Line.Visible = msoFalse
        anchor.Fill.Visible = msoTrue
        anchor.Fill.Transparency = 1.0
        group = slide.Shapes.Range([duplicate.Name, anchor.Name]).Group()
        # Use the actual group bounds, including objects outside the slide.
        origin_left, origin_top = group.Left, group.Top
        bounds_width, bounds_height = group.Width, group.Height
        # Render at a positive origin, then map back to the original coordinates.
        group.Left = 0
        group.Top = 0
        group.Export(temp_png_path, ppShapeFormatPNG,
                     round(slide_width_pt * 2), round(slide_height_pt * 2))
        cropped = crop_transparent_area(temp_png_path, bounds_width, bounds_height)
        if not cropped:
            raise RuntimeError("No visible pixels found")
        left, top, width, height = cropped
        new_shape = slide.Shapes.AddPicture(
            temp_png_path, False, True, origin_left + left, origin_top + top, width, height
        )
        new_shape.Name = shape.Name + "_img"
        # Rotation is already rendered into the PNG; do not rotate it again.
        group.Delete()
        group = duplicate = anchor = None
        while new_shape.ZOrderPosition > original_z:
            new_shape.ZOrder(msoSendBackward)
        # Commit last: every operation above can fail without changing the original.
        shape.Delete()
        return True
    except Exception:
        cleanup_errors = []
        objects = [new_shape, group] if group is not None else [new_shape, duplicate, anchor]
        for obj in objects:
            if obj is not None:
                try:
                    obj.Delete()
                except Exception as exc:
                    cleanup_errors.append(str(exc))
        if cleanup_errors:
            raise ConversionRecoveryError("Temporary shape cleanup failed; output was not saved")
        raise

def text_to_image_ppt(input_ppt, progress_callback=None, texts=None):
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
    original_alerts = original_security = None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    output_ppt = "%s_converted_%s%s" % (base_name, timestamp, ext)
    abs_output_path = os.path.abspath(output_ppt)

    try:
        temp_fd, temp_png_path = tempfile.mkstemp(prefix="ppt_text_", suffix=".png")
        os.close(temp_fd)

        try:
            try:
                powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
                owns_powerpoint = True
        except Exception as exc:
            logging.warning("PowerPoint launch failed: %s", exc)
            raise make_powerpoint_required_error(texts)
        original_alerts = powerpoint.DisplayAlerts
        original_security = powerpoint.AutomationSecurity
        for opened in powerpoint.Presentations:
            if os.path.normcase(os.path.abspath(opened.FullName)) == os.path.normcase(abs_input_path):
                raise UserVisibleError(texts["dialog_file_locked_title"], texts["error_file_locked"],
                                       kind="file_open")
        powerpoint.DisplayAlerts = 1
        powerpoint.AutomationSecurity = 3  # Disable macros in programmatically opened files.

        try:
            presentation = powerpoint.Presentations.Open(abs_input_path, ReadOnly=True, Untitled=True, WithWindow=False)
        except Exception as exc:
            logging.warning("Presentation open failed: %s", exc)
            raise classify_open_error(exc, texts)
        finally:
            powerpoint.AutomationSecurity = original_security

        total_slides = presentation.Slides.Count
        slide_width_pt = presentation.PageSetup.SlideWidth
        slide_height_pt = presentation.PageSetup.SlideHeight

        if progress_callback:
            progress_callback(0, total_slides, texts["status_start"])

        result = ConversionResult(abs_output_path)

        for slide_index, slide in enumerate(presentation.Slides, start=1):
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
            ungroup_all_shapes(slide, result.issues, slide_index)

            shapes = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]
            for i, shape in enumerate(shapes, start=1):
                try:
                    if shape.HasTable or shape.HasChart or shape.HasSmartArt:
                        result.issues.append({"slide": slide_index, "shape_index": i,
                                              "stage": "unsupported",
                                              "error": "Table, chart or SmartArt requires manual review"})
                        continue

                    if not has_visible_text(shape):
                        continue

                    ok = shape_to_cropped_picture(
                        slide=slide,
                        shape=shape,
                        slide_width_pt=slide_width_pt,
                        slide_height_pt=slide_height_pt,
                        temp_png_path=temp_png_path
                    )

                    if ok:
                        result.converted_count += 1

                except ConversionRecoveryError:
                    raise
                except Exception as e:
                    result.issues.append({"slide": slide_index, "shape_index": i,
                                          "stage": "convert", "error": str(e)})
                    logging.warning(
                        "Slide %s shape %s skipped: %s",
                        slide_index,
                        i,
                        e
                    )

            if progress_callback:
                progress_callback(
                    slide_index,
                    total_slides,
                    texts["status_slide_done"].format(
                        current=slide_index,
                        total=total_slides
                    )
                )

        try:
            presentation.SaveAs(abs_output_path, 1 if ext.lower() == ".ppt" else 24)
        except Exception as exc:
            logging.warning("Presentation save failed: %s", exc)
            raise classify_save_error(exc, texts)

        result.report_path = abs_output_path + ".report.json"
        try:
            with open(result.report_path, "w", encoding="utf-8") as report:
                json.dump({"converted_count": result.converted_count, "issues": result.issues},
                          report, ensure_ascii=False, indent=2)
        except OSError as exc:
            result.report_path = ""
            result.issues.append({"stage": "report", "error": str(exc)})

        if progress_callback:
            progress_callback(
                total_slides,
                total_slides,
                texts["status_done_count"].format(count=result.converted_count)
            )

        return result

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass

        if powerpoint is not None:
            try:
                if original_alerts is not None:
                    powerpoint.DisplayAlerts = original_alerts
                if original_security is not None:
                    powerpoint.AutomationSecurity = original_security
                if owns_powerpoint and powerpoint.Presentations.Count == 0:
                    powerpoint.Quit()
            except Exception as exc:
                logging.warning("PowerPoint cleanup failed: %s", exc)

        if temp_png_path and os.path.exists(temp_png_path):
            try:
                os.remove(temp_png_path)
            except Exception:
                pass

        pythoncom.CoUninitialize()

        
class RoundedPanel(tk.Canvas):
    """Rounded surface containing regular, keyboard-accessible Tk widgets."""

    def __init__(self, parent, background, surface, border, **kwargs):
        super().__init__(parent, bg=background, highlightthickness=0, bd=0, height=1, **kwargs)
        self.surface = surface
        self.border = border
        self.body = tk.Frame(self, bg=surface, padx=6, pady=2)
        self.surface_id = self.create_polygon(0, 0, 0, 0, fill=surface, outline=border, smooth=True)
        self.body_id = self.create_window(16, 16, window=self.body, anchor="nw")
        self.bind("<Configure>", self.layout)
        self.body.bind("<Configure>", self.layout)

    def layout(self, event=None):
        width = self.winfo_width()
        if width <= 32:
            return
        self.itemconfigure(self.body_id, width=max(1, width - 32))
        height = self.body.winfo_reqheight() + 32
        if int(self.cget("height")) != height:
            self.configure(height=height)
        radius = 18
        self.coords(self.surface_id, radius, 1, width-radius, 1, width-1, 1,
                    width-1, radius, width-1, height-radius, width-1, height-1,
                    width-radius, height-1, radius, height-1, 1, height-1,
                    1, height-radius, 1, radius, 1, 1)


class App:
    WINDOW_WIDTH = 1000
    BASE_WINDOW_HEIGHT = 730
    BG_COLOR = "#F5F7FB"
    CARD_COLOR = "#FFFFFF"
    BORDER_COLOR = "#E2E8F1"
    PRIMARY_COLOR = "#5258E8"
    PRIMARY_DARK = "#4147C7"
    TEXT_COLOR = "#182238"
    MUTED_COLOR = "#69758A"
    WARNING_BG = "#FFF8E9"
    WARNING_TEXT = "#85621C"
    SUCCESS_COLOR = "#18856B"
    ERROR_COLOR = "#C84051"
    SIDEBAR_COLOR = "#151E33"

    def __init__(self, root):
        self.root = root
        self.lang_code = "ko"
        self.selected_file = ""
        self.is_converting = False
        self.worker_thread = None
        self.ui_event_queue = queue.Queue()
        self.status_key = "status_ready"
        self.root.geometry(f"{self.WINDOW_WIDTH}x{self.BASE_WINDOW_HEIGHT}")
        self.root.minsize(940, 650)
        self.root.resizable(True, True)
        self.root.configure(bg=self.BG_COLOR)
        self.root.option_add("*Font", ("Malgun Gothic", 10))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Primary.TButton", font=("Malgun Gothic", 11, "bold"),
                        padding=(22, 13), background=self.PRIMARY_COLOR, foreground="white",
                        borderwidth=0, focusthickness=2, focuscolor="#A9ADF8")
        style.map("Primary.TButton", background=[("disabled", "#DFE3F0"),
                  ("pressed", self.PRIMARY_DARK), ("active", "#6268F0")],
                  foreground=[("disabled", "#7C879C")])
        style.configure("Secondary.TButton", font=("Malgun Gothic", 10, "bold"),
                        padding=(18, 9), background="#EEF0FF", foreground=self.PRIMARY_COLOR,
                        borderwidth=0, focusthickness=2, focuscolor="#B9BDF7")
        style.map("Secondary.TButton", background=[("active", "#E2E5FF")],
                  foreground=[("disabled", "#8D96AD")])
        style.configure("Language.TCombobox", padding=(8, 6), fieldbackground="white",
                        background="white", bordercolor=self.BORDER_COLOR,
                        arrowcolor=self.MUTED_COLOR, foreground=self.TEXT_COLOR)
        style.map("Language.TCombobox", fieldbackground=[("readonly", "white")],
                  selectbackground=[("readonly", "white")],
                  selectforeground=[("readonly", self.TEXT_COLOR)])
        style.configure("App.Horizontal.TProgressbar", thickness=6, troughcolor="#EBEEF5",
                        background=self.PRIMARY_COLOR, borderwidth=0, lightcolor=self.PRIMARY_COLOR,
                        darkcolor=self.PRIMARY_COLOR)
        self.progress_style = style

        self.sidebar = tk.Frame(root, bg=self.SIDEBAR_COLOR, width=254, padx=26, pady=30)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        brand = tk.Frame(self.sidebar, bg=self.SIDEBAR_COLOR)
        brand.pack(fill="x")
        tk.Label(brand, text="P", font=("Segoe UI", 15, "bold"), bg=self.PRIMARY_COLOR,
                 fg="white", padx=10, pady=3).pack(side="left")
        tk.Label(brand, text="PPT / PIXEL", bg=self.SIDEBAR_COLOR, fg="#F3F5FF",
                 font=("Segoe UI", 11, "bold"), padx=12).pack(side="left")
        tk.Label(self.sidebar, text="Aa", bg=self.SIDEBAR_COLOR, fg="#E1E3FF",
                 font=("Segoe UI", 74, "bold"), anchor="w").pack(fill="x", pady=(42, 0))
        tk.Label(self.sidebar, text="TEXT   →   PNG", bg=self.SIDEBAR_COLOR,
                 fg="#A4A9FF", font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")
        self.lbl_side_title = tk.Label(self.sidebar, bg=self.SIDEBAR_COLOR, fg="white",
                                      font=("Malgun Gothic", 15, "bold"), justify="left", anchor="w")
        self.lbl_side_title.pack(fill="x", pady=(28, 10))
        self.lbl_side_body = tk.Label(self.sidebar, bg=self.SIDEBAR_COLOR, fg="#9DA9C1",
                                     font=("Malgun Gothic", 9), justify="left", anchor="w", wraplength=200)
        self.lbl_side_body.pack(fill="x")
        tk.Frame(self.sidebar, bg="#2A344B", height=1).pack(fill="x", pady=28)
        self.step_labels = []
        for number in (1, 2, 3):
            row = tk.Frame(self.sidebar, bg=self.SIDEBAR_COLOR)
            row.pack(fill="x", pady=7)
            tk.Label(row, text=f"0{number}", bg=self.SIDEBAR_COLOR, fg="#8D95BD",
                     font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 12))
            label = tk.Label(row, bg=self.SIDEBAR_COLOR, fg="#CBD2E2", font=("Malgun Gothic", 9))
            label.pack(side="left")
            self.step_labels.append(label)
        self.lbl_local = tk.Label(self.sidebar, bg=self.SIDEBAR_COLOR, fg="#9DA9C1",
                                 font=("Malgun Gothic", 9), anchor="w")
        self.lbl_local.pack(side="bottom", fill="x")

        workspace = tk.Frame(root, bg=self.BG_COLOR)
        workspace.pack(side="right", fill="both", expand=True)
        workspace.rowconfigure(0, weight=1)
        workspace.columnconfigure(0, weight=1)
        self.content_canvas = tk.Canvas(workspace, bg=self.BG_COLOR, highlightthickness=0)
        self.content_canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(workspace, orient="vertical", command=self.content_canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.content_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.container = tk.Frame(self.content_canvas, bg=self.BG_COLOR, padx=34, pady=22)
        self.content_id = self.content_canvas.create_window(0, 0, window=self.container, anchor="nw")
        self.content_canvas.bind("<Configure>", self.sync_scroll)
        self.container.bind("<Configure>", self.sync_scroll)
        self.root.bind("<MouseWheel>", self.scroll_content)
        toolbar = tk.Frame(self.container, bg=self.BG_COLOR)
        toolbar.pack(fill="x")
        self.lbl_workspace = tk.Label(toolbar, bg=self.BG_COLOR, fg=self.MUTED_COLOR,
                                     font=("Malgun Gothic", 9, "bold"))
        self.lbl_workspace.pack(side="left")
        self.cmb_lang = ttk.Combobox(toolbar, state="readonly", width=9,
                                    values=["한국어", "English"], style="Language.TCombobox")
        self.cmb_lang.current(0)
        self.cmb_lang.bind("<<ComboboxSelected>>", self.change_language)
        self.cmb_lang.pack(side="right")
        self.lbl_title = tk.Label(self.container, bg=self.BG_COLOR, fg=self.TEXT_COLOR,
                                 font=("Malgun Gothic", 25, "bold"), anchor="w", justify="left")
        self.lbl_title.pack(fill="x", pady=(20, 8))
        self.lbl_subtitle = tk.Label(self.container, bg=self.BG_COLOR, fg=self.MUTED_COLOR,
                                    font=("Malgun Gothic", 10), anchor="w", justify="left", wraplength=600)
        self.lbl_subtitle.pack(fill="x", pady=(0, 20))

        self.file_panel = RoundedPanel(self.container, self.BG_COLOR, self.CARD_COLOR, self.BORDER_COLOR)
        self.file_panel.pack(fill="x")
        file_header = tk.Frame(self.file_panel.body, bg=self.CARD_COLOR)
        file_header.pack(fill="x")
        tk.Label(file_header, text="PPT", bg="#EEF0FF", fg=self.PRIMARY_COLOR,
                 font=("Segoe UI", 10, "bold"), padx=10, pady=8).pack(side="left", padx=(0, 14))
        file_text = tk.Frame(file_header, bg=self.CARD_COLOR)
        file_text.pack(side="left", fill="x", expand=True)
        self.lbl_file_title = tk.Label(file_text, bg=self.CARD_COLOR, fg=self.TEXT_COLOR,
                                      font=("Malgun Gothic", 11, "bold"), anchor="w")
        self.lbl_file_title.pack(fill="x")
        self.lbl_formats = tk.Label(file_text, bg=self.CARD_COLOR, fg=self.MUTED_COLOR,
                                   font=("Malgun Gothic", 9), anchor="w")
        self.lbl_formats.pack(fill="x", pady=(3, 0))
        self.lbl_file_value = tk.Label(self.file_panel.body, bg=self.CARD_COLOR, fg=self.MUTED_COLOR,
                                      font=("Malgun Gothic", 9), anchor="w", justify="left", wraplength=550)
        self.lbl_file_value.pack(fill="x", pady=(14, 12))
        self.btn_select = ttk.Button(self.file_panel.body, command=self.select_file, style="Secondary.TButton")
        self.btn_select.pack(anchor="w")

        self.progress_panel = RoundedPanel(self.container, self.BG_COLOR, self.CARD_COLOR, self.BORDER_COLOR)
        self.progress_panel.pack(fill="x", pady=(12, 0))
        progress_header = tk.Frame(self.progress_panel.body, bg=self.CARD_COLOR)
        progress_header.pack(fill="x")
        self.lbl_progress_title = tk.Label(progress_header, bg=self.CARD_COLOR, fg=self.TEXT_COLOR,
                                          font=("Malgun Gothic", 10, "bold"))
        self.lbl_progress_title.pack(side="left")
        self.lbl_percent = tk.Label(progress_header, text="0%", bg=self.CARD_COLOR, fg=self.PRIMARY_COLOR,
                                    font=("Segoe UI", 12, "bold"))
        self.lbl_percent.pack(side="right")
        self.progress_bar = ttk.Progressbar(self.progress_panel.body, orient="horizontal",
                                            mode="determinate", style="App.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(10, 8))
        self.lbl_status = tk.Label(self.progress_panel.body, bg=self.CARD_COLOR, fg=self.MUTED_COLOR,
                                  font=("Malgun Gothic", 9), anchor="w", justify="left", wraplength=550)
        self.lbl_status.pack(fill="x")

        self.warning_frame = tk.Frame(self.container, bg=self.BG_COLOR)
        self.warning_frame.pack(fill="x", pady=(16, 16))
        tk.Label(self.warning_frame, text="i", bg=self.BG_COLOR, fg=self.MUTED_COLOR,
                 font=("Segoe UI", 11, "bold")).pack(side="left", anchor="n", padx=(0, 10))
        self.lbl_warning_body = tk.Label(self.warning_frame, bg=self.BG_COLOR, fg=self.MUTED_COLOR,
                                        font=("Malgun Gothic", 9), justify="left", anchor="w", wraplength=550)
        self.lbl_warning_body.pack(side="left", fill="x", expand=True)
        self.btn_convert = ttk.Button(self.container, command=self.start_conversion, style="Primary.TButton")
        self.btn_convert.pack(fill="x")
        self.lbl_output_hint = tk.Label(self.container, bg=self.BG_COLOR, fg=self.MUTED_COLOR,
                                       font=("Malgun Gothic", 9), anchor="center", wraplength=580)
        self.lbl_output_hint.pack(fill="x", pady=(10, 0))
        self.lbl_requirements = tk.Label(self.container, bg=self.BG_COLOR, fg="#8490A4",
                                        font=("Malgun Gothic", 8), anchor="center")
        self.lbl_requirements.pack(side="bottom", fill="x", pady=(10, 0))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Control-o>", lambda event: self.select_file() if not self.is_converting else None)
        self.root.bind("<Configure>", self.on_resize)
        self.apply_language()

    def on_resize(self, event):
        if event.widget is self.root:
            width = max(400, event.width - 254 - 68)
            for label in (self.lbl_subtitle, self.lbl_file_value, self.lbl_status, self.lbl_warning_body):
                label.configure(wraplength=width - 44)

    def sync_scroll(self, event=None):
        width = self.content_canvas.winfo_width()
        self.content_canvas.itemconfigure(self.content_id, width=width)
        required = self.container.winfo_reqheight()
        available = self.content_canvas.winfo_height()
        self.content_canvas.itemconfigure(self.content_id, height=max(required, available))
        self.content_canvas.configure(scrollregion=(0, 0, width, max(required, available)))
        if required > available:
            self.scrollbar.grid()
        else:
            self.scrollbar.grid_remove()
            self.content_canvas.yview_moveto(0)

    def scroll_content(self, event):
        if self.container.winfo_reqheight() > self.content_canvas.winfo_height():
            self.content_canvas.yview_scroll(-int(event.delta / 120), "units")

    def on_close(self):
        if self.is_converting:
            messagebox.showwarning(self.tr("app_title"), self.tr("dialog_close_busy"))
            return
        self.root.destroy()

    def tr(self, key, **kwargs):
        text = LANG[self.lang_code][key]
        if kwargs:
            return text.format(**kwargs)
        return text

    def apply_language(self):
        self.root.title(self.tr("app_title"))
        for label, key in ((self.lbl_title, "main_title"), (self.lbl_subtitle, "subtitle"),
                           (self.lbl_workspace, "workspace"), (self.lbl_side_title, "side_title"),
                           (self.lbl_side_body, "side_body"), (self.lbl_local, "local_note"),
                           (self.lbl_formats, "file_formats"), (self.lbl_output_hint, "output_hint"),
                           (self.lbl_warning_body, "compact_warning"),
                           (self.lbl_progress_title, "progress_title"),
                           (self.lbl_requirements, "requirements")):
            label.config(text=self.tr(key))
        self.lbl_requirements.config(text=f'{self.tr("requirements")} · v{APP_VERSION}')
        for label, key in zip(self.step_labels, ("step_select", "step_convert", "step_save")):
            label.config(text=self.tr(key))
        self.btn_convert.config(text=self.tr("btn_convert"))
        self.update_selected_file_label()
        if self.status_key in ("status_ready", "ready_hint", "selected_hint"):
            self.status_key = "selected_hint" if self.selected_file else "ready_hint"
        self.lbl_status.config(text=self.tr(self.status_key))
        self.fit_window_height()

    def change_language(self, event=None):
        selected = self.cmb_lang.get()
        self.lang_code = "en" if selected == "English" else "ko"
        self.apply_language()

    def format_display_path(self, path, max_chars=90):
        if len(path) <= max_chars:
            return path

        head_len = (max_chars - 3) // 2
        tail_len = max_chars - 3 - head_len
        return "%s...%s" % (path[:head_len], path[-tail_len:])

    def update_selected_file_label(self):
        selected = bool(self.selected_file)
        self.lbl_file_title.config(text=self.tr("file_label" if selected else "file_heading"))
        self.lbl_file_value.config(
            text=self.format_display_path(self.selected_file, 110) if selected else self.tr("file_placeholder"),
            fg=self.TEXT_COLOR if selected else self.MUTED_COLOR
        )
        self.btn_select.config(text=self.tr("file_change" if selected else "btn_select"))
        self.btn_convert.config(state="normal" if selected and not self.is_converting else "disabled")
        self.fit_window_height()

    def fit_window_height(self):
        self.root.update_idletasks()

    def set_busy(self, busy):
        self.is_converting = busy
        self.btn_select.config(state="disabled" if busy else "normal")
        self.btn_convert.config(state="disabled" if busy or not self.selected_file else "normal",
                                text=self.tr("btn_busy" if busy else "btn_convert"))
        self.cmb_lang.config(state="disabled" if busy else "readonly")
        self.root.config(cursor="watch" if busy else "")
        self.root.update_idletasks()

    def set_status(self, text, color=None):
        self.lbl_status.config(text=text, fg=color or self.MUTED_COLOR)
        self.root.update_idletasks()

    def update_progress(self, current, total, message):
        self.progress_bar["maximum"] = max(total, 1)
        self.progress_bar["value"] = current
        self.lbl_percent.config(text=f"{min(100, round(current / max(total, 1) * 100))}%")
        self.set_status(message, self.PRIMARY_COLOR)

    def enqueue_progress(self, current, total, message):
        self.ui_event_queue.put({
            "type": "progress",
            "current": current,
            "total": total,
            "message": message
        })

    def run_conversion_worker(self, filepath, texts):
        try:
            result = text_to_image_ppt(
                filepath,
                progress_callback=self.enqueue_progress,
                texts=texts
            )
            self.ui_event_queue.put({
                "type": "done",
                "output_path": result.output_path,
                "converted_count": result.converted_count,
                "issues": result.issues,
                "report_path": result.report_path
            })
        except UserVisibleError as exc:
            logging.warning("User visible error [%s]: %s", exc.kind, exc.message)
            self.ui_event_queue.put({
                "type": "error",
                "kind": exc.kind,
                "title": exc.title,
                "error": exc.message
            })
        except Exception as exc:
            logging.exception("Conversion failed")
            self.ui_event_queue.put({
                "type": "error",
                "kind": "generic",
                "error": str(exc)
            })

    def process_ui_events(self):
        while True:
            try:
                event = self.ui_event_queue.get_nowait()
            except queue.Empty:
                break

            event_type = event["type"]

            if event_type == "progress":
                self.update_progress(
                    event["current"],
                    event["total"],
                    event["message"]
                )
            elif event_type == "done":
                self.progress_bar["value"] = self.progress_bar["maximum"]
                self.lbl_percent.config(text="100%")
                self.set_busy(False)
                if event["issues"]:
                    self.status_key = "status_partial"
                    self.set_status(self.tr("status_partial"), self.WARNING_TEXT)
                    messagebox.showwarning(
                        self.tr("dialog_partial_title"),
                        self.tr("dialog_partial_msg", count=event["converted_count"],
                                failed=len(event["issues"]), path=event["output_path"],
                                report=event["report_path"] or str(event["issues"]))
                    )
                    continue
                self.status_key = "status_done"
                self.set_status(self.tr("status_done"), self.SUCCESS_COLOR)
                messagebox.showinfo(
                    self.tr("dialog_done_title"),
                    self.tr(
                        "dialog_done_msg",
                        count=event["converted_count"],
                        path=event["output_path"]
                    )
                )
            elif event_type == "error":
                self.progress_bar["value"] = 0
                self.lbl_percent.config(text="0%")
                self.status_key = "status_error"
                self.set_status(self.tr("status_error"), self.ERROR_COLOR)
                self.set_busy(False)
                if event.get("title"):
                    messagebox.showerror(
                        event["title"],
                        event["error"]
                    )
                else:
                    messagebox.showerror(
                        self.tr("dialog_error_title"),
                        self.tr("dialog_error_msg", error=event["error"])
                    )

        if self.is_converting or not self.ui_event_queue.empty():
            self.root.after(100, self.process_ui_events)

    def select_file(self):
        if self.is_converting:
            return
        filepath = filedialog.askopenfilename(
            title=self.tr("dialog_select"),
            filetypes=(("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*"))
        )

        if not filepath:
            return

        self.selected_file = filepath
        self.update_selected_file_label()
        self.status_key = "selected_hint"
        self.update_progress(0, 1, self.tr("selected_hint"))

    def start_conversion(self):
        if self.is_converting or not self.selected_file:
            return
        filepath = self.selected_file
        confirmed = messagebox.askokcancel(
            self.tr("dialog_confirm_title"),
            self.tr(
                "dialog_confirm_msg",
                name=os.path.basename(filepath)
            )
        )

        if not confirmed:
            self.status_key = "selected_hint"
            self.set_status(self.tr("selected_hint"))
            return

        self.set_busy(True)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 1
        self.lbl_percent.config(text="0%")
        self.status_key = "status_preparing"
        self.set_status(self.tr("status_preparing"), self.PRIMARY_COLOR)
        texts = LANG[self.lang_code].copy()
        self.worker_thread = threading.Thread(
            target=self.run_conversion_worker,
            args=(filepath, texts),
            daemon=False
        )
        self.worker_thread.start()
        self.root.after(100, self.process_ui_events)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
