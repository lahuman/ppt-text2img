import os
import queue
import logging
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, simpledialog, ttk

from runtime_support import configure_logging
from translations import LANG
# Keep the existing app-level conversion API available to callers.
from conversion import (
    ConversionResult, ConversionCancelled, ConversionRecoveryError,
    UserVisibleError, PowerPointRequiredError, text_to_image_ppt,
    shape_to_cropped_picture, crop_transparent_area, has_visible_text,
    ungroup_all_shapes, ensure_runtime_requirements,
)

APP_VERSION = "1.4"

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


class FontChoiceDialog(simpledialog.Dialog):
    def __init__(self, parent, faces, texts):
        self.faces = faces
        self.texts = texts
        super().__init__(parent, texts["font_title"])

    def body(self, master):
        ttk.Label(master, text=self.texts["font_notice"], wraplength=470).pack(anchor="w", pady=6)
        details = tk.Text(master, width=54, height=min(9, 2 + 2 * len(self.faces)), wrap="word")
        for name, face in self.faces.items():
            numbers = ", ".join(map(str, sorted(face.slides))) or self.texts["font_not_direct"]
            details.insert("end", f"{name} · {self.texts['font_slides']} {numbers}\n")
            if face.shared:
                numbers = ", ".join(map(str, sorted(face.related_slides)))
                details.insert("end", self.texts["font_shared"] + (" " + numbers if numbers else "") + "\n")
        details.configure(state="disabled")
        details.pack(fill="x")
        ttk.Label(master, text=self.texts["font_replacement"]).pack(anchor="w", pady=(10, 2))
        installed = set(tkfont.families(master))
        groups = [("맑은 고딕", "Malgun Gothic"), ("Arial",), ("Segoe UI",)]
        restricted = {face.casefold() for face in self.faces}
        choices = [next(f for f in group if f in installed) for group in groups
                   if installed.intersection(group) and not restricted.intersection(f.casefold() for f in group)]
        self.selector = ttk.Combobox(master, values=choices, state="readonly")
        if choices:
            self.selector.current(0)
        self.selector.pack(fill="x")
        return self.selector

    def buttonbox(self):
        bar = ttk.Frame(self)
        bar.pack(pady=10)
        ttk.Button(bar, text=self.texts["font_apply"], command=self.ok,
                   state="normal" if self.selector.get() else "disabled").pack(side="left", padx=4)
        ttk.Button(bar, text=self.texts["font_cancel"], command=self.cancel).pack(side="left")
        self.bind("<Escape>", self.cancel)

    def apply(self):
        self.result = self.selector.get()


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
        if os.name == "nt":
            root.iconbitmap(default=os.path.join(os.path.dirname(__file__), "icon.ico"))
        self.lang_code = "ko"
        self.selected_file = ""
        self.is_converting = False
        self.result_event = None
        self.saving = False
        self.worker_thread = None
        self.cancel_event = threading.Event()
        self.started_at = None
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
                                  font=("Malgun Gothic", 12, "bold"), anchor="w", justify="left", wraplength=550)
        self.lbl_status.pack(fill="x")

        self.result_actions = tk.Frame(self.progress_panel.body, bg=self.CARD_COLOR)
        self.result_actions.pack(fill="x")
        self.btn_output = ttk.Button(self.result_actions, style="Primary.TButton")
        self.btn_report = ttk.Button(self.result_actions, style="Secondary.TButton")
        self.warning_frame = tk.Frame(self.container, bg=self.BG_COLOR)
        self.warning_frame.pack(fill="x", pady=(16, 16))
        info = tk.Canvas(self.warning_frame, width=20, height=20, bg=self.BG_COLOR,
                         highlightthickness=0)
        info.pack(side="left", anchor="n", padx=(0, 8))
        info.create_oval(2, 2, 18, 18, outline=self.PRIMARY_COLOR)
        info.create_oval(9, 6, 11, 8, fill=self.PRIMARY_COLOR, outline="")
        info.create_line(10, 10, 10, 14, fill=self.PRIMARY_COLOR, width=2)
        self.lbl_warning_body = tk.Label(self.warning_frame, bg=self.BG_COLOR, fg=self.MUTED_COLOR,
                                        font=("Malgun Gothic", 9), justify="left", anchor="w", wraplength=550)
        self.lbl_warning_body.pack(side="left", fill="x", expand=True)
        self.btn_convert = ttk.Button(self.container, command=self.start_conversion, style="Primary.TButton")
        self.btn_convert.pack(fill="x")
        self.btn_cancel = ttk.Button(self.container, command=self.request_cancel, style="Secondary.TButton")
        self.btn_cancel.pack(fill="x", pady=(6, 0))
        self.lbl_elapsed = tk.Label(self.container, bg=self.BG_COLOR, fg=self.MUTED_COLOR)
        self.lbl_elapsed.pack(fill="x")
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
        self.btn_cancel.config(text=self.tr("btn_cancel"), state="disabled")
        self.update_selected_file_label()
        if self.status_key in ("status_ready", "ready_hint", "selected_hint"):
            self.status_key = "selected_hint" if self.selected_file else "ready_hint"
        self.lbl_status.config(text=self.tr(self.status_key))
        if self.result_event:
            self.show_result(self.result_event)
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
        self.btn_cancel.config(state="normal" if busy else "disabled")
        self.btn_select.config(state="disabled" if busy else "normal")
        self.btn_convert.config(state="disabled" if busy or not self.selected_file else "normal",
                                text=self.tr("btn_busy" if busy else "btn_convert"))
        self.cmb_lang.config(state="disabled" if busy else "readonly")
        if busy:
            self.file_panel.pack_forget()
            self.btn_convert.pack_forget()
            self.lbl_progress_title.config(text=self.format_display_path(os.path.basename(self.selected_file), 52))
            self.lbl_warning_body.config(text=self.tr("minimize_hint"))
        else:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
            self.file_panel.pack(before=self.progress_panel, fill="x")
            self.btn_convert.pack(before=self.btn_cancel, fill="x")
            self.lbl_progress_title.config(text=self.tr("progress_title"))
            self.lbl_warning_body.config(text=self.tr("compact_warning"))
        self.root.update_idletasks()

    def set_status(self, text, color=None):
        self.lbl_status.config(text=text, fg=color or self.MUTED_COLOR)
        self.root.update_idletasks()

    def update_progress(self, current, total, message):
        if str(self.progress_bar["mode"]) == "indeterminate":
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
        self.progress_bar["maximum"] = max(total, 1)
        self.progress_bar["value"] = current
        self.lbl_percent.config(text=f"{min(99 if self.is_converting else 100, round(current / max(total, 1) * 100))}%")
        self.set_status(message, self.PRIMARY_COLOR)

    def enqueue_progress(self, current, total, message):
        self.ui_event_queue.put({
            "type": "progress",
            "current": current,
            "total": total,
            "message": message
        })

    def request_font_choice(self, faces):
        reply = queue.Queue(maxsize=1)
        self.ui_event_queue.put({"type": "font_choice", "faces": faces, "reply": reply})
        while not self.cancel_event.is_set():
            try:
                return reply.get(timeout=0.1)
            except queue.Empty:
                continue
        return None

    def run_conversion_worker(self, filepath, texts):
        try:
            result = text_to_image_ppt(
                filepath,
                progress_callback=self.enqueue_progress,
                font_choice=self.request_font_choice,
                saving_callback=lambda: self.ui_event_queue.put({"type": "saving"}),
                texts=texts,
                cancel_event=self.cancel_event
            )
            self.ui_event_queue.put({
                "type": "done",
                "state": result.state,
                "output_path": result.output_path,
                "converted_count": result.converted_count,
                "issues": result.issues
            })
        except ConversionCancelled:
            self.ui_event_queue.put({"type": "cancelled"})
        except UserVisibleError as exc:
            logging.warning("User visible error [%s]", exc.kind)
            self.ui_event_queue.put({
                "type": "error",
                "kind": exc.kind,
                "title": exc.title,
                "error": exc.message
            })
        except Exception as exc:
            logging.error("Conversion failed (%s)", type(exc).__name__)
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

            if event_type == "font_choice":
                choice = None
                try:
                    if not self.cancel_event.is_set():
                        choice = FontChoiceDialog(self.root, event["faces"], LANG[self.lang_code]).result
                finally:
                    event["reply"].put(choice)
            elif event_type == "saving":
                self.saving = True
                self.btn_cancel.config(state="disabled")
                self.show_activity("status_saving")
            elif event_type == "progress":
                if self.saving or self.cancel_event.is_set():
                    continue
                self.update_progress(
                    event["current"],
                    event["total"],
                    event["message"]
                )
            elif event_type == "done":
                self.set_busy(False)
                self.progress_bar["value"] = self.progress_bar["maximum"]
                self.lbl_percent.config(text="100%" if event["state"] != "recovery_required" else "")
                self.show_result(event)
            elif event_type == "cancelled":
                self.set_busy(False)
                self.progress_bar["value"] = 0
                self.lbl_percent.config(text="0%")
                self.status_key = "status_cancelled"
                self.set_status(self.tr("status_cancelled"))
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

        if self.is_converting and self.started_at is not None:
            self.lbl_elapsed.config(text=self.tr("elapsed", seconds=int(time.monotonic() - self.started_at)))
        if self.is_converting or not self.ui_event_queue.empty():
            self.root.after(100, self.process_ui_events)

    def show_activity(self, key):
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start(20)
        self.lbl_percent.config(text="")
        self.set_status(self.tr(key), self.PRIMARY_COLOR)

    def open_path(self, path):
        try:
            os.startfile(path)
        except OSError:
            messagebox.showerror(self.tr("dialog_error_title"), self.tr("open_failed", path=path))

    def clear_result(self):
        self.result_event = None
        self.btn_output.pack_forget()
        self.btn_report.pack_forget()
        self.lbl_elapsed.config(text="")

    def show_result(self, event):
        self.result_event = event
        path = event["output_path"]
        state = event["state"]
        recovery = state == "recovery_required"
        self.status_key = {"complete": "status_done", "partial": "status_partial",
                           "recovery_required": "status_recovery_required"}[state]
        warnings = sum(issue.get("stage") == "unsupported" for issue in event["issues"])
        message = self.tr("result_counts", count=event["converted_count"], warnings=warnings,
                          errors=len(event["issues"]) - warnings)
        if recovery:
            message = self.tr("status_recovery_required") + "\n" + message
        self.set_status(message, self.SUCCESS_COLOR if state == "complete" else self.WARNING_TEXT)
        target = os.path.dirname(path) if recovery else path
        self.btn_output.config(text=self.tr("open_recovery" if recovery else "open_result"),
                               command=lambda: self.open_path(target))
        self.btn_output.pack(side="left", pady=(10, 0))
        self.btn_output.configure(style="Primary.TButton")
        self.btn_report.pack_forget()
        if event["issues"]:
            self.btn_report.config(text=self.tr("open_report"), command=lambda: self.save_report(event))
            self.btn_report.pack(side="left", padx=8, pady=(10, 0))

    def save_report(self, event):
        output = Path(event["output_path"])
        path = filedialog.asksaveasfilename(parent=self.root, title=self.tr("open_report"),
                                            initialdir=str(output.parent), initialfile=output.stem + ".txt",
                                            defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if not path:
            return
        if Path(path).suffix.lower() != ".txt":
            messagebox.showerror(self.tr("open_report"), self.tr("report_txt"), parent=self.root)
            return
        lines = [output.name, self.tr("report_hint"), ""]
        for issue in event["issues"]:
            stage = issue.get("stage", "")
            location = self.tr("report_document")
            if "slide" in issue:
                location = f"{self.tr('issue_slide')} {issue['slide']}"
                if issue.get("slide_title"):
                    location += " — " + issue["slide_title"]
            if issue.get("shape_name"):
                location += f" · {self.tr('issue_shape_index')} «{issue['shape_name']}»"
            if issue.get("excerpt"):
                location += f" · “{issue['excerpt']}”"
            reason = stage if stage in ("unsupported", "ungroup", "convert") else "cleanup"
            lines.append(location)
            lines.append(self.tr("report_" + reason))
            if stage != "unsupported":
                lines.append(f"{stage} — {issue.get('error', '')}")
            lines.append("")
        try:
            Path(path).write_text("\n".join(lines), encoding="utf-8-sig")
        except OSError as exc:
            messagebox.showerror(self.tr("open_report"), str(exc), parent=self.root)
            return
        messagebox.showinfo(self.tr("open_report"), self.tr("report_saved", path=path), parent=self.root)

    def select_file(self):
        if self.is_converting:
            return
        filepath = filedialog.askopenfilename(
            title=self.tr("dialog_select"),
            filetypes=(("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*"))
        )

        if not filepath:
            return

        self.clear_result()
        self.selected_file = filepath
        self.update_selected_file_label()
        self.status_key = "selected_hint"
        self.update_progress(0, 1, self.tr("selected_hint"))

    def request_cancel(self):
        if self.is_converting:
            self.cancel_event.set()
            self.btn_cancel.config(state="disabled")
            self.status_key = "status_cancelling"
            self.set_status(self.tr("status_cancelling"))

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

        self.clear_result()
        self.saving = False
        self.cancel_event.clear()
        self.started_at = time.monotonic()
        self.lbl_elapsed.config(text="")
        self.set_busy(True)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 1
        self.lbl_percent.config(text="0%")
        self.status_key = "status_preparing"
        self.show_activity("opening")
        texts = LANG[self.lang_code].copy()
        self.worker_thread = threading.Thread(
            target=self.run_conversion_worker,
            args=(filepath, texts),
            daemon=False
        )
        try:
            self.worker_thread.start()
        except RuntimeError as exc:
            self.set_busy(False)
            self.status_key = "status_error"
            self.set_status(self.tr("status_error"), self.ERROR_COLOR)
            messagebox.showerror(self.tr("dialog_error_title"),
                                 self.tr("dialog_error_msg", error=str(exc)))
            return
        self.root.after(100, self.process_ui_events)


if __name__ == "__main__":
    configure_logging()
    root = tk.Tk()
    app = App(root)
    root.mainloop()
