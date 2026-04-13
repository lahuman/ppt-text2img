import os
import tempfile
import logging
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image

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
        "main_title": "PPT 텍스트를 이미지로 안전하게 변환합니다",
        "subtitle": "배포용 PPT에서 폰트 깨짐을 막기 위해 텍스트를 이미지로 바꿉니다.",
        "lang_label": "언어",
        "btn_select": "PPT 파일 선택하기",
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
        "dialog_error_title": "에러",
        "dialog_error_msg": "변환 중 오류가 발생했습니다:\n{error}",
        "error_windows_only": "이 프로그램은 Windows에서만 실행할 수 있습니다.",
        "error_powerpoint_required": "Microsoft PowerPoint와 pywin32가 설치된 Windows 환경이 필요합니다.\n상세 오류: {error}",
        "lang_ko": "한국어",
        "lang_en": "English",
    },
    "en": {
        "app_title": "PPT Text to Image Converter",
        "main_title": "Convert PPT text into images safely",
        "subtitle": "Turn text into images to prevent font and layout issues on other PCs.",
        "lang_label": "Language",
        "btn_select": "Select PPT File",
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
        "dialog_error_title": "Error",
        "dialog_error_msg": "An error occurred during conversion:\n{error}",
        "error_windows_only": "This program can only run on Windows.",
        "error_powerpoint_required": "A Windows environment with Microsoft PowerPoint and pywin32 is required.\nDetails: {error}",
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


def has_visible_text(shape):
    try:
        if not shape.HasTextFrame:
            return False
        if not shape.TextFrame.HasText:
            return False

        text = shape.TextFrame.TextRange.Text
        return bool(text and text.strip())
    except Exception:
        return False


def ungroup_all_shapes(slide):
    changed = True
    while changed:
        changed = False
        for idx in range(slide.Shapes.Count, 0, -1):
            try:
                shp = slide.Shapes(idx)
                if shp.Type == MsoShapeTypeGroup:
                    shp.Ungroup()
                    changed = True
                    break
            except Exception as e:
                logging.debug("Ungroup skipped at index %s: %s", idx, e)


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
        raise RuntimeError(texts["error_windows_only"])

    if pythoncom is None or win32com is None:
        raise RuntimeError(
            texts["error_powerpoint_required"].format(error=COM_IMPORT_ERROR)
        )


def shape_to_cropped_picture(slide, shape, slide_width_pt, slide_height_pt, temp_png_path):
    orig_name = ""
    orig_rot = 0.0
    anchor = None
    group = None
    anchor_name = ""

    try:
        try:
            orig_name = shape.Name
        except Exception:
            orig_name = ""

        try:
            orig_rot = shape.Rotation
            shape.Rotation = 0
        except Exception:
            orig_rot = 0.0

        # 투명 anchor 추가
        anchor = slide.Shapes.AddShape(1, 0, 0, slide_width_pt, slide_height_pt)
        try:
            anchor_name = "__ppt_anchor_%d" % anchor.Id
            anchor.Name = anchor_name
        except Exception:
            try:
                anchor_name = anchor.Name
            except Exception:
                anchor_name = ""

        anchor.Line.Visible = 0
        anchor.Fill.Visible = -1
        anchor.Fill.Transparency = 1.0

        # 원본 shape + anchor 를 직접 그룹화
        # Duplicate() 사용 안 함
        group = slide.Shapes.Range([orig_name, anchor_name]).Group()
        group.Export(temp_png_path, ppShapeFormatPNG)

        cropped = crop_transparent_area(temp_png_path, slide_width_pt, slide_height_pt)
        if not cropped:
            raise RuntimeError("No visible pixels found")

        left_pt, top_pt, width_pt, height_pt = cropped

        new_shape = slide.Shapes.AddPicture(
            temp_png_path,
            False,
            True,
            left_pt,
            top_pt,
            width_pt,
            height_pt
        )

        try:
            if orig_name:
                new_shape.Name = orig_name + "_img"
        except Exception:
            pass

        try:
            new_shape.Rotation = orig_rot
        except Exception:
            pass

        # 핵심: 원본을 따로 Delete 하지 말고
        # 원본이 포함된 group 을 통째로 삭제
        try:
            group.Delete()
            group = None
        except Exception as e:
            raise RuntimeError("Group delete failed: %s" % e)

        return True

    except Exception as e:
        logging.warning("Shape conversion skipped: %s", e)

        # 실패 시 원복 시도
        try:
            if group is not None:
                group.Ungroup()
        except Exception:
            pass

        if anchor_name:
            try:
                slide.Shapes(anchor_name).Delete()
            except Exception:
                pass

        if orig_name:
            try:
                slide.Shapes(orig_name).Rotation = orig_rot
            except Exception:
                pass

        return False

def text_to_image_ppt(input_ppt, progress_callback=None, texts=None):
    texts = texts or LANG["ko"]
    ensure_runtime_requirements(texts)

    pythoncom.CoInitialize()

    powerpoint = None
    presentation = None
    temp_png_path = None

    base_name, ext = os.path.splitext(input_ppt)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_ppt = "%s_converted_%s%s" % (base_name, timestamp, ext)
    abs_input_path = os.path.abspath(input_ppt)
    abs_output_path = os.path.abspath(output_ppt)

    try:
        temp_fd, temp_png_path = tempfile.mkstemp(prefix="ppt_text_", suffix=".png")
        os.close(temp_fd)

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.DisplayAlerts = 0

        presentation = powerpoint.Presentations.Open(abs_input_path, WithWindow=False)

        total_slides = presentation.Slides.Count
        slide_width_pt = presentation.PageSetup.SlideWidth
        slide_height_pt = presentation.PageSetup.SlideHeight

        if progress_callback:
            progress_callback(0, total_slides, texts["status_start"])

        converted_count = 0

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

            # 그룹 내부 텍스트도 빠짐없이 변환되도록 먼저 모두 해제한다.
            ungroup_all_shapes(slide)

            for i in range(slide.Shapes.Count, 0, -1):
                try:
                    shape = slide.Shapes(i)

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
                        converted_count += 1

                except Exception as e:
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

        presentation.SaveAs(abs_output_path)

        if progress_callback:
            progress_callback(
                total_slides,
                total_slides,
                texts["status_done_count"].format(count=converted_count)
            )

        return abs_output_path, converted_count

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass

        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass

        if temp_png_path and os.path.exists(temp_png_path):
            try:
                os.remove(temp_png_path)
            except Exception:
                pass

        pythoncom.CoUninitialize()

        
class App:
    BG_COLOR = "#F4F1EA"
    CARD_COLOR = "#FFFDF8"
    BORDER_COLOR = "#DDD4C5"
    PRIMARY_COLOR = "#1F6F78"
    PRIMARY_DARK = "#15565E"
    TEXT_COLOR = "#1F2933"
    MUTED_COLOR = "#667085"
    WARNING_BG = "#FFF6DB"
    WARNING_TEXT = "#7A5A00"
    SUCCESS_COLOR = "#1A7F37"
    ERROR_COLOR = "#B42318"

    def __init__(self, root):
        self.root = root
        self.lang_code = "ko"
        self.selected_file = ""

        self.root.geometry("640x560")
        self.root.resizable(False, False)
        self.root.configure(bg=self.BG_COLOR)

        self.progress_style = ttk.Style()
        if "clam" in self.progress_style.theme_names():
            self.progress_style.theme_use("clam")
        self.progress_style.configure(
            "App.Horizontal.TProgressbar",
            thickness=14,
            troughcolor="#E9E1D5",
            background=self.PRIMARY_COLOR,
            bordercolor="#E9E1D5",
            lightcolor=self.PRIMARY_COLOR,
            darkcolor=self.PRIMARY_COLOR
        )

        self.container = tk.Frame(
            root,
            bg=self.BG_COLOR,
            padx=24,
            pady=22
        )
        self.container.pack(fill="both", expand=True)

        self.header_frame = tk.Frame(self.container, bg=self.BG_COLOR)
        self.header_frame.pack(fill="x")

        self.title_frame = tk.Frame(self.header_frame, bg=self.BG_COLOR)
        self.title_frame.pack(side="left", fill="x", expand=True)

        self.lbl_title = tk.Label(
            self.title_frame,
            bg=self.BG_COLOR,
            fg=self.TEXT_COLOR,
            font=("Malgun Gothic", 18, "bold"),
            justify="left",
            anchor="w"
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(
            self.title_frame,
            bg=self.BG_COLOR,
            fg=self.MUTED_COLOR,
            font=("Malgun Gothic", 10),
            justify="left",
            anchor="w",
            wraplength=430
        )
        self.lbl_subtitle.pack(anchor="w", pady=(6, 0))

        self.lang_frame = tk.Frame(self.header_frame, bg=self.BG_COLOR)
        self.lang_frame.pack(side="right", anchor="n", padx=(16, 0))

        self.lbl_lang = tk.Label(
            self.lang_frame,
            bg=self.BG_COLOR,
            fg=self.MUTED_COLOR,
            font=("Malgun Gothic", 10, "bold")
        )
        self.lbl_lang.pack(anchor="e")

        self.cmb_lang = ttk.Combobox(
            self.lang_frame,
            state="readonly",
            width=10,
            values=["한국어", "English"]
        )
        self.cmb_lang.current(0)
        self.cmb_lang.bind("<<ComboboxSelected>>", self.change_language)
        self.cmb_lang.pack(anchor="e", pady=(6, 0))

        self.info_card = tk.Frame(
            self.container,
            bg=self.CARD_COLOR,
            highlightbackground=self.BORDER_COLOR,
            highlightthickness=1,
            padx=18,
            pady=18
        )
        self.info_card.pack(fill="x", pady=(20, 14))

        self.lbl_file_title = tk.Label(
            self.info_card,
            bg=self.CARD_COLOR,
            fg=self.MUTED_COLOR,
            font=("Malgun Gothic", 10, "bold"),
            anchor="w"
        )
        self.lbl_file_title.pack(fill="x")

        self.lbl_file_value = tk.Label(
            self.info_card,
            bg=self.CARD_COLOR,
            fg=self.TEXT_COLOR,
            font=("Malgun Gothic", 10),
            anchor="w",
            justify="left",
            wraplength=540
        )
        self.lbl_file_value.pack(fill="x", pady=(8, 10))

        self.lbl_output_hint = tk.Label(
            self.info_card,
            bg=self.CARD_COLOR,
            fg=self.MUTED_COLOR,
            font=("Malgun Gothic", 9),
            anchor="w",
            justify="left",
            wraplength=540
        )
        self.lbl_output_hint.pack(fill="x")

        self.btn_select = tk.Button(
            self.info_card,
            command=self.select_file,
            bg=self.PRIMARY_COLOR,
            fg="white",
            activebackground=self.PRIMARY_DARK,
            activeforeground="white",
            disabledforeground="#D8EEEF",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Malgun Gothic", 11, "bold"),
            width=24,
            height=2
        )
        self.btn_select.pack(anchor="w", pady=(16, 0))

        self.warning_frame = tk.LabelFrame(
            self.container,
            bg=self.WARNING_BG,
            fg=self.WARNING_TEXT,
            bd=1,
            relief="solid",
            padx=16,
            pady=12,
            font=("Malgun Gothic", 10, "bold")
        )
        self.warning_frame.pack(fill="x", pady=(0, 14))

        self.lbl_warning_body = tk.Label(
            self.warning_frame,
            bg=self.WARNING_BG,
            fg=self.WARNING_TEXT,
            font=("Malgun Gothic", 10),
            anchor="w",
            justify="left",
            wraplength=560
        )
        self.lbl_warning_body.pack(fill="x")

        self.progress_card = tk.Frame(
            self.container,
            bg=self.CARD_COLOR,
            highlightbackground=self.BORDER_COLOR,
            highlightthickness=1,
            padx=18,
            pady=18
        )
        self.progress_card.pack(fill="x")

        self.lbl_progress_title = tk.Label(
            self.progress_card,
            bg=self.CARD_COLOR,
            fg=self.MUTED_COLOR,
            font=("Malgun Gothic", 10, "bold"),
            anchor="w"
        )
        self.lbl_progress_title.pack(fill="x")

        self.lbl_status = tk.Label(
            self.progress_card,
            bg=self.CARD_COLOR,
            fg=self.PRIMARY_COLOR,
            font=("Malgun Gothic", 10),
            anchor="w",
            justify="left",
            wraplength=540
        )
        self.lbl_status.pack(fill="x", pady=(8, 12))

        self.progress_bar = ttk.Progressbar(
            self.progress_card,
            orient="horizontal",
            length=540,
            mode="determinate",
            style="App.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill="x")

        self.apply_language()

    def tr(self, key, **kwargs):
        text = LANG[self.lang_code][key]
        if kwargs:
            return text.format(**kwargs)
        return text

    def apply_language(self):
        self.root.title(self.tr("app_title"))
        self.lbl_lang.config(text=self.tr("lang_label"))
        self.lbl_title.config(text=self.tr("main_title"))
        self.lbl_subtitle.config(text=self.tr("subtitle"))
        self.lbl_file_title.config(text=self.tr("file_label"))
        self.lbl_output_hint.config(text=self.tr("output_hint"))
        self.btn_select.config(text=self.tr("btn_select"))
        self.warning_frame.config(text=self.tr("warning_title"))
        self.lbl_warning_body.config(text=self.tr("warning_body"))
        self.lbl_progress_title.config(text=self.tr("progress_title"))
        self.update_selected_file_label()
        self.lbl_status.config(text=self.tr("status_ready"), fg=self.PRIMARY_COLOR)

    def change_language(self, event=None):
        selected = self.cmb_lang.get()
        self.lang_code = "en" if selected == "English" else "ko"
        self.apply_language()

    def update_selected_file_label(self):
        if self.selected_file:
            self.lbl_file_value.config(text=self.selected_file, fg=self.TEXT_COLOR)
        else:
            self.lbl_file_value.config(
                text=self.tr("file_placeholder"),
                fg=self.MUTED_COLOR
            )

    def set_busy(self, busy):
        self.btn_select.config(
            state="disabled" if busy else "normal",
            text=self.tr("btn_busy") if busy else self.tr("btn_select"),
            cursor="watch" if busy else "hand2"
        )
        self.cmb_lang.config(state="disabled" if busy else "readonly")
        self.root.config(cursor="watch" if busy else "")
        self.root.update_idletasks()

    def set_status(self, text, color="blue"):
        self.lbl_status.config(text=text, fg=color)
        self.root.update_idletasks()

    def update_progress(self, current, total, message):
        self.progress_bar["maximum"] = max(total, 1)
        self.progress_bar["value"] = current
        self.set_status(message, self.PRIMARY_COLOR)

    def select_file(self):
        filepath = filedialog.askopenfilename(
            title=self.tr("dialog_select"),
            filetypes=(("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*"))
        )

        if not filepath:
            return

        self.selected_file = filepath
        self.update_selected_file_label()

        confirmed = messagebox.askokcancel(
            self.tr("dialog_confirm_title"),
            self.tr(
                "dialog_confirm_msg",
                name=os.path.basename(filepath)
            )
        )

        if not confirmed:
            self.set_status(self.tr("status_ready"), self.PRIMARY_COLOR)
            return

        self.set_busy(True)
        self.progress_bar["value"] = 0
        self.set_status(self.tr("status_preparing"), self.PRIMARY_COLOR)

        try:
            output_path, converted_count = text_to_image_ppt(
                filepath,
                progress_callback=self.update_progress,
                texts=LANG[self.lang_code]
            )

            self.set_status(self.tr("status_done"), self.SUCCESS_COLOR)
            messagebox.showinfo(
                self.tr("dialog_done_title"),
                self.tr(
                    "dialog_done_msg",
                    count=converted_count,
                    path=output_path
                )
            )

        except Exception as e:
            logging.exception("Conversion failed")
            self.set_status(self.tr("status_error"), self.ERROR_COLOR)
            messagebox.showerror(
                self.tr("dialog_error_title"),
                self.tr("dialog_error_msg", error=e)
            )

        finally:
            self.progress_bar["value"] = 0
            self.set_busy(False)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
