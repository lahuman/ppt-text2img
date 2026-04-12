import os
import tempfile
import logging
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client
from PIL import Image

LANG = {
    "ko": {
        "app_title": "PPT 글씨 -> 이미지 변환기",
        "main_title": "PPT 파일의 텍스트를 이미지로 변환합니다.\n(폰트 깨짐 방지용)",
        "lang_label": "언어",
        "btn_select": "PPT 파일 선택하기",
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
        "dialog_error_title": "에러",
        "dialog_error_msg": "변환 중 오류가 발생했습니다:\n{error}",
        "lang_ko": "한국어",
        "lang_en": "English",
    },
    "en": {
        "app_title": "PPT Text to Image Converter",
        "main_title": "Convert all text in a PPT file\ninto images.\n(Prevents font corruption)",
        "lang_label": "Language",
        "btn_select": "Select PPT File",
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
        "dialog_error_title": "Error",
        "dialog_error_msg": "An error occurred during conversion:\n{error}",
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
        group.Export(temp_png_path, 2)  # ppShapeFormatPNG

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

            # 필요 시 사용
            # ungroup_all_shapes(slide)

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
    def __init__(self, root):
        self.root = root
        self.lang_code = "ko"

        self.root.geometry("500x320")
        self.root.resizable(False, False)

        # 언어 선택 영역
        self.lang_frame = tk.Frame(root)
        self.lang_frame.pack(pady=(12, 4))

        self.lbl_lang = tk.Label(self.lang_frame, font=("Arial", 10))
        self.lbl_lang.pack(side="left", padx=(0, 8))

        self.cmb_lang = ttk.Combobox(
            self.lang_frame,
            state="readonly",
            width=10,
            values=["한국어", "English"]
        )
        self.cmb_lang.current(0)
        self.cmb_lang.bind("<<ComboboxSelected>>", self.change_language)
        self.cmb_lang.pack(side="left")

        self.lbl_title = tk.Label(
            root,
            font=("Arial", 12),
            justify="center"
        )
        self.lbl_title.pack(pady=16)

        self.btn_select = tk.Button(
            root,
            command=self.select_file,
            font=("Arial", 10),
            width=24,
            height=2
        )
        self.btn_select.pack(pady=6)

        self.lbl_status = tk.Label(
            root,
            fg="blue",
            font=("Arial", 10),
            width=55,
            anchor="w",
            justify="left",
            wraplength=430
        )
        self.lbl_status.pack(pady=8)

        self.progress_bar = ttk.Progressbar(
            root,
            orient="horizontal",
            length=380,
            mode="determinate"
        )
        self.progress_bar.pack(pady=6)

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
        self.btn_select.config(text=self.tr("btn_select"))
        self.lbl_status.config(text=self.tr("status_ready"), fg="blue")

    def change_language(self, event=None):
        selected = self.cmb_lang.get()
        self.lang_code = "en" if selected == "English" else "ko"
        self.apply_language()

    def set_status(self, text, color="blue"):
        self.lbl_status.config(text=text, fg=color)
        self.root.update_idletasks()

    def update_progress(self, current, total, message):
        self.progress_bar["maximum"] = max(total, 1)
        self.progress_bar["value"] = current
        self.set_status(message, "blue")

    def select_file(self):
        filepath = filedialog.askopenfilename(
            title=self.tr("dialog_select"),
            filetypes=(("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*"))
        )

        if not filepath:
            return

        self.btn_select.config(state="disabled")
        self.cmb_lang.config(state="disabled")
        self.progress_bar["value"] = 0
        self.set_status(self.tr("status_preparing"), "blue")

        try:
            output_path, converted_count = text_to_image_ppt(
                filepath,
                progress_callback=self.update_progress,
                texts=LANG[self.lang_code]
            )

            self.set_status(self.tr("status_done"), "green")
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
            self.set_status(self.tr("status_error"), "red")
            messagebox.showerror(
                self.tr("dialog_error_title"),
                self.tr("dialog_error_msg", error=e)
            )

        finally:
            self.progress_bar["value"] = 0
            self.btn_select.config(state="normal")
            self.cmb_lang.config(state="readonly")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()