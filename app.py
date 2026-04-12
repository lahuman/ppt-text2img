import win32com.client
import os
from datetime import datetime
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk 

def text_to_image_ppt(input_ppt):
    base_name, ext = os.path.splitext(input_ppt)
    timestamp = datetime.now().strftime("%H%M%S")
    output_ppt = f"{base_name}_converted_{timestamp}{ext}"
    
    powerpoint = None
    
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        abs_input_path = os.path.abspath(input_ppt)
        presentation = powerpoint.Presentations.Open(abs_input_path, WithWindow=False)

        temp_img_path = os.path.join(tempfile.gettempdir(), "temp_ppt_shape.png")
        
        total_slides = presentation.Slides.Count
        progress_bar["maximum"] = total_slides
        progress_bar["value"] = 0

        # 슬라이드 크기 (EMU → 포인트 변환: 914400 EMU = 1인치 = 72pt)
        # PowerPoint COM은 포인트 단위 사용
        slide_width  = presentation.PageSetup.SlideWidth   # pt
        slide_height = presentation.PageSetup.SlideHeight  # pt

        for slide_index, slide in enumerate(presentation.Slides, 1):
            
            lbl_status.config(text=f"변환 중... (슬라이드 {slide_index} / {total_slides} 완료)")
            progress_bar["value"] = slide_index
            root.update()

            # 1. 모든 그룹 해제
            has_group = True
            while has_group:
                has_group = False
                for i in range(slide.Shapes.Count, 0, -1):
                    try:
                        shape = slide.Shapes(i)
                        if shape.Type == 6:
                            shape.Ungroup()
                            has_group = True
                            break
                    except:
                        pass

            # 2. 텍스트 → 이미지 변환
            for i in range(slide.Shapes.Count, 0, -1):
                shape = slide.Shapes(i)
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    try:
                        orig_rot = shape.Rotation
                        shape.Rotation = 0

                        # 원본 텍스트 박스의 위치/크기 저장
                        orig_left   = shape.Left
                        orig_top    = shape.Top
                        orig_width  = shape.Width
                        orig_height = shape.Height

                        if orig_width <= 0 or orig_height <= 0:
                            continue

                        # ── 투명 보호막 사각형 생성 ──────────────────────────────
                        # 슬라이드 전체 크기로 만들어, 그룹 바운딩박스가
                        # 항상 슬라이드 전체와 같아지도록 고정합니다.
                        # 이렇게 하면 Export된 이미지에서의
                        # "원본 텍스트 박스 위치 비율"이 정확하게 계산됩니다.
                        anchor = slide.Shapes.AddShape(
                            1,               # msoShapeRectangle
                            0, 0,            # 슬라이드 좌상단에 고정
                            slide_width,
                            slide_height
                        )
                        anchor.Line.Visible = 0
                        anchor.Fill.Visible = -1
                        anchor.Fill.Transparency = 1.0  # 완전 투명

                        # 텍스트 도형 + 앵커 사각형을 그룹으로 묶기
                        group = slide.Shapes.Range([i, slide.Shapes.Count]).Group()

                        # 그룹은 항상 슬라이드 크기(0,0 ~ slide_width, slide_height)
                        g_left   = group.Left    # == 0
                        g_top    = group.Top     # == 0
                        g_width  = group.Width   # == slide_width
                        g_height = group.Height  # == slide_height

                        # 슬라이드 전체 크기로 이미지 Export
                        group.Export(temp_img_path, 2)  # 2 = ppShapeFormatPNG

                        # ── 이미지에서 원본 도형의 비율 계산 ────────────────────
                        # Export된 이미지는 (g_width x g_height) 비율로 렌더링됨.
                        # 원본 도형이 이미지 내에서 차지하는 비율을 구해,
                        # 슬라이드에 다시 원본 크기로 정확히 배치합니다.

                        # 이미지 내 원본 도형의 픽셀 위치 비율
                        ratio_x = (orig_left - g_left) / g_width
                        ratio_y = (orig_top  - g_top)  / g_height
                        ratio_w = orig_width  / g_width
                        ratio_h = orig_height / g_height

                        # 이미지 전체를 슬라이드 크기로 삽입한 뒤 크롭하는 방식 대신,
                        # 이미지를 원본 도형 크기에 맞게 "역산"하여 삽입합니다.
                        # 즉, 이미지를 원본 크기의 역수로 확대하면
                        # 텍스트 부분이 정확히 원본 도형 영역을 채웁니다.

                        # 삽입할 이미지의 전체 크기 (원본 도형 크기 / 비율)
                        insert_width  = orig_width  / ratio_w   # == slide_width
                        insert_height = orig_height / ratio_h   # == slide_height

                        # 삽입 위치: 이미지의 (ratio_x, ratio_y) 지점이
                        # 원본 도형의 (orig_left, orig_top)에 오도록 역산
                        insert_left = orig_left - ratio_x * insert_width
                        insert_top  = orig_top  - ratio_y * insert_height

                        new_shape = slide.Shapes.AddPicture(
                            temp_img_path,
                            False, True,
                            insert_left,
                            insert_top,
                            insert_width,
                            insert_height
                        )

                        # 회전각 재적용
                        new_shape.Rotation = orig_rot

                        # 원본 그룹 삭제
                        group.Delete()

                    except Exception as inner_e:
                        print(f"도형 변환 무시됨: {inner_e}")
                        pass

        abs_output_path = os.path.abspath(output_ppt)
        presentation.SaveAs(abs_output_path)
        presentation.Close()
        powerpoint.Quit()

        if os.path.exists(temp_img_path):
            try: os.remove(temp_img_path)
            except: pass

        lbl_status.config(text="변환이 모두 완료되었습니다!", fg="green")
        messagebox.showinfo("완료", f"변환이 완료되었습니다!\n저장 위치: {output_ppt}")
        progress_bar["value"] = 0
        btn_select.config(state="normal")

    except Exception as e:
        lbl_status.config(text="오류 발생!", fg="red")
        messagebox.showerror("에러", f"변환 중 오류가 발생했습니다:\n{str(e)}")
        progress_bar["value"] = 0
        btn_select.config(state="normal")
        if powerpoint:
            try: powerpoint.Quit()
            except: pass


def select_file():
    filepath = filedialog.askopenfilename(
        title="변환할 PPT 파일을 선택하세요",
        filetypes=(("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*"))
    )
    if filepath:
        btn_select.config(state="disabled")
        lbl_status.config(text="프로그램을 준비 중입니다...", fg="blue")
        root.update()
        text_to_image_ppt(filepath)


# --- GUI ---
root = tk.Tk()
root.title("PPT 글씨 -> 이미지 변환기")
root.geometry("400x250")

lbl_title = tk.Label(root, text="PPT 파일의 모든 텍스트를\n이미지로 변환합니다.", font=("Arial", 12))
lbl_title.pack(pady=20)

btn_select = tk.Button(root, text="PPT 파일 선택하기", command=select_file, font=("Arial", 10), width=20, height=2)
btn_select.pack(pady=5)

lbl_status = tk.Label(root, text="대기 중", fg="blue", font=("Arial", 10))
lbl_status.pack(pady=5)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=5)

root.mainloop()
