# -*- coding: utf-8 -*-
"""
图片文字提取工具
================
- 基于 RapidOCR 的本地离线免费 OCR 引擎，无需联网
- 支持 JPG / PNG / BMP / WEBP 图片
- 选择图片 -> 开始识别 -> 一键复制文字
- 隐藏自检模式: main.py --selftest <图片路径> [输出文件路径]
"""

import os
import queue
import sys
import threading
import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageOps

try:
    import openpyxl
    from openpyxl.styles import Font, Border, Side, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

APP_TITLE = "图片文字提取工具"
APP_SUBTITLE = "本地离线识别 · 免费 · 支持 JPG / PNG"
SUPPORTED_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".webp")

# ---------------------------------------------------------------------------
# OCR 引擎（全局单例，懒加载，首次识别时初始化）
# ---------------------------------------------------------------------------
_ocr_engine = None
_engine_lock = threading.Lock()


def get_ocr_engine():
    global _ocr_engine
    if _ocr_engine is None:
        with _engine_lock:
            if _ocr_engine is None:
                from rapidocr_onnxruntime import RapidOCR
                _ocr_engine = RapidOCR()
    return _ocr_engine


def recognize_image(path):
    """识别图片中的文字，返回按原图顺序拼接的文本。"""
    engine = get_ocr_engine()
    out = engine(path)
    if out is None:
        return ""
    # 兼容 rapidocr 新旧版本返回格式
    if isinstance(out, tuple):
        result = out[0] or []
        lines = []
        for item in result:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                lines.append(str(item[1]))
        return "\n".join(lines)
    txts = getattr(out, "txts", None)
    if txts:
        return "\n".join(txts)
    return ""


def recognize_with_boxes(path):
    """识别图片中的文字，返回 [(text, bbox), ...]，bbox 为四个角点坐标。"""
    engine = get_ocr_engine()
    out = engine(path)
    if out is None:
        return []
    items = []
    if isinstance(out, tuple):
        result = out[0] or []
        for item in result:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                items.append((str(item[1]), item[0]))
    else:
        txts = getattr(out, "txts", None)
        boxes = getattr(out, "boxes", None)
        if txts and boxes:
            for t, b in zip(txts, boxes):
                items.append((str(t), b))
    return items


def extract_table(ocr_items):
    """
    根据 OCR 结果的坐标，把文字按行列分组，还原为二维表格。
    ocr_items: [(text, bbox), ...]，bbox 为 [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
    返回: list[list[str]]
    """
    if not ocr_items:
        return []

    parsed = []
    for text, bbox in ocr_items:
        try:
            xs = [float(p[0]) for p in bbox]
            ys = [float(p[1]) for p in bbox]
        except (TypeError, ValueError, IndexError):
            continue
        cx = sum(xs) / 4.0
        cy = sum(ys) / 4.0
        left = min(xs)
        right = max(xs)
        top = min(ys)
        bottom = max(ys)
        height = max(bottom - top, 1.0)
        parsed.append((text, cx, cy, left, right, height))

    if not parsed:
        return []

    # 按 y 中心排序，按行聚类
    parsed.sort(key=lambda x: x[2])
    rows = []
    current_row = [parsed[0]]
    for item in parsed[1:]:
        prev_cy = current_row[-1][2]
        threshold = item[5] * 0.55  # 行高的 55% 作为同行判定阈值
        if abs(item[2] - prev_cy) <= threshold:
            current_row.append(item)
        else:
            rows.append(current_row)
            current_row = [item]
    rows.append(current_row)

    # 每行内按左边界排序，提取文字
    table = []
    for row in rows:
        row.sort(key=lambda x: x[3])
        table.append([item[0] for item in row])

    return table


def export_table_to_excel(table_data, save_path):
    """把二维表格数据写入 xlsx 文件，带表头加粗、边框、居中、自动列宽。"""
    if not HAS_OPENPYXL:
        raise RuntimeError("未安装 openpyxl，无法导出 Excel")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表格提取结果"

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, row in enumerate(table_data, 1):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            cell.alignment = center
            if r == 1:
                cell.font = Font(bold=True)

    # 自动列宽
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            try:
                v = str(cell.value) if cell.value is not None else ""
                # 中文按 2 个宽度估算
                w = sum(2 if ord(ch) > 127 else 1 for ch in v)
                if w > max_len:
                    max_len = w
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 10), 50)

    wb.save(save_path)
    return save_path


# ---------------------------------------------------------------------------
# 主界面
# ---------------------------------------------------------------------------
class OCRApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1040x700")
        self.minsize(900, 620)

        self.image_path = None
        self._preview_ctk_image = None
        self._recognizing = False
        self._msg_queue = queue.Queue()

        self._set_icon()
        self._build_ui()
        self.after(100, self._poll_queue)

    # ---------------- UI 构建 ----------------
    def _set_icon(self):
        try:
            self.iconbitmap(resource_path("app_icon.ico"))
        except Exception:
            pass

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 顶栏
        header = ctk.CTkFrame(self, corner_radius=0, fg_color="#121318", height=66)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header, text=APP_TITLE,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold"),
            text_color="#FFFFFF",
        ).grid(row=0, column=0, sticky="w", padx=24, pady=14)
        ctk.CTkLabel(
            header, text=APP_SUBTITLE,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color="#8A8F9C",
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        header.grid_columnconfigure(2, weight=1)

        # 主题切换开关
        self.theme_switch = ctk.CTkSwitch(
            header, text="深色模式", width=90,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            command=self.toggle_theme,
        )
        self.theme_switch.grid(row=0, column=3, sticky="e", padx=(0, 20))
        self.theme_switch.select()  # 默认深色

        # 主体
        body = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=(14, 6))
        body.grid_columnconfigure(0, weight=1, uniform="col")
        body.grid_columnconfigure(1, weight=1, uniform="col")
        body.grid_rowconfigure(0, weight=1)

        # ---- 左：图片预览 ----
        left = ctk.CTkFrame(body, corner_radius=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)

        self.preview_label = ctk.CTkLabel(
            left,
            text="尚未选择图片\n\n点击下方按钮，从电脑中选择 JPG / PNG 图片",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=15),
            text_color="#7A808C",
        )
        self.preview_label.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)

        self.select_btn = ctk.CTkButton(
            left, text="选择图片", height=44,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=15, weight="bold"),
            corner_radius=10, command=self.select_image,
        )
        self.select_btn.grid(row=1, column=0, sticky="ew", padx=16, pady=14)

        # ---- 右：识别结果 ----
        right = ctk.CTkFrame(body, corner_radius=16)
        right.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        result_header = ctk.CTkFrame(right, fg_color="transparent")
        result_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))
        ctk.CTkLabel(
            result_header, text="识别结果",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=15, weight="bold"),
        ).pack(side="left")
        self.char_count_label = ctk.CTkLabel(
            result_header, text="",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color="#7A808C",
        )
        self.char_count_label.pack(side="left", padx=(10, 0))

        self.result_box = ctk.CTkTextbox(
            right, wrap="word",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=15),
            corner_radius=10,
        )
        self.result_box.grid(row=1, column=0, sticky="nsew", padx=16, pady=(4, 8))

        btn_row = ctk.CTkFrame(right, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 14))
        btn_row.grid_columnconfigure(0, weight=3)
        btn_row.grid_columnconfigure(1, weight=3)
        btn_row.grid_columnconfigure(2, weight=2)
        btn_row.grid_columnconfigure(3, weight=2)

        self.recognize_btn = ctk.CTkButton(
            btn_row, text="开始识别", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            corner_radius=10, command=self.start_recognize,
        )
        self.recognize_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.table_btn = ctk.CTkButton(
            btn_row, text="表格提取", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            corner_radius=10, fg_color="#1565C0", hover_color="#1976D2",
            command=self.start_table_extract,
        )
        self.table_btn.grid(row=0, column=1, sticky="ew", padx=6)

        self.copy_btn = ctk.CTkButton(
            btn_row, text="复制文字", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            corner_radius=10, fg_color="#2E7D32", hover_color="#388E3C",
            state="disabled", command=self.copy_text,
        )
        self.copy_btn.grid(row=0, column=2, sticky="ew", padx=6)

        self.clear_btn = ctk.CTkButton(
            btn_row, text="清空", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            corner_radius=10, fg_color="#3A3F4B", hover_color="#4A5060",
            command=self.clear_all,
        )
        self.clear_btn.grid(row=0, column=3, sticky="ew", padx=(6, 0))

        # ---- 底部状态栏 ----
        footer = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 12))
        footer.grid_columnconfigure(0, weight=1)
        self.status_label = ctk.CTkLabel(
            footer, text="就绪：请选择一张图片开始",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color="#8A8F9C",
        )
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress_bar = ctk.CTkProgressBar(
            footer, width=180, height=8, mode="indeterminate"
        )
        self.progress_bar.grid(row=0, column=1, sticky="e")
        self.progress_bar.set(0)

    # ---------------- 交互逻辑 ----------------
    def toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("dark")
            self.theme_switch.configure(text="深色模式")
        else:
            ctk.set_appearance_mode("light")
            self.theme_switch.configure(text="浅色模式")

    def select_image(self):
        if self._recognizing:
            return
        path = filedialog.askopenfilename(
            title="选择要识别的图片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.webp"), ("所有文件", "*.*")],
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in SUPPORTED_EXT:
            messagebox.showwarning("格式不支持", "请选择 JPG 或 PNG 格式的图片。")
            return
        try:
            self._update_preview(path)
        except Exception as exc:
            messagebox.showerror("无法打开图片", f"图片文件可能已损坏或不是有效图片：\n{exc}")
            return
        self.image_path = path
        self.set_status(f"已选择：{os.path.basename(path)}，点击「开始识别」")

    def _update_preview(self, path):
        img = Image.open(path)
        img = ImageOps.exif_transpose(img)
        img = img.convert("RGB")
        img.thumbnail((470, 420), Image.LANCZOS)
        self._preview_ctk_image = ctk.CTkImage(
            light_image=img, dark_image=img, size=img.size
        )
        self.preview_label.configure(image=self._preview_ctk_image, text="")

    def start_recognize(self):
        if self._recognizing:
            return
        if not self.image_path:
            messagebox.showinfo("提示", "请先选择一张图片。")
            return
        self._recognizing = True
        self.select_btn.configure(state="disabled")
        self.recognize_btn.configure(state="disabled", text="识别中…")
        self.table_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.result_box.delete("1.0", "end")
        self.char_count_label.configure(text="")
        self.progress_bar.start()
        self.set_status("正在识别文字，请稍候…")
        threading.Thread(
            target=self._recognize_worker, args=(self.image_path,), daemon=True
        ).start()

    def _recognize_worker(self, path):
        try:
            text = recognize_image(path)
            self._msg_queue.put(("done", text))
        except Exception as exc:
            self._msg_queue.put(("error", str(exc)))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self._msg_queue.get_nowait()
                if kind == "done":
                    self._on_done(payload)
                elif kind == "error":
                    self._on_error(payload)
                elif kind == "table_done":
                    self._on_table_done(payload)
                elif kind == "table_empty":
                    self._on_table_empty()
                elif kind == "table_error":
                    self._on_table_error(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _on_done(self, text):
        self._recognizing = False
        self.select_btn.configure(state="normal")
        self.recognize_btn.configure(state="normal", text="开始识别")
        self.table_btn.configure(state="normal")
        self.progress_bar.stop()
        self.result_box.insert("1.0", text)
        if text.strip():
            n = len(text.replace("\n", "").replace(" ", ""))
            self.char_count_label.configure(text=f"共 {n} 个字符")
            self.copy_btn.configure(state="normal")
            self.set_status(f"识别完成，共 {n} 个字符，可直接复制")
        else:
            self.set_status("未在图片中识别到文字，可尝试更清晰的图片")

    def _on_error(self, err):
        self._recognizing = False
        self.select_btn.configure(state="normal")
        self.recognize_btn.configure(state="normal", text="开始识别")
        self.table_btn.configure(state="normal")
        self.progress_bar.stop()
        self.set_status("识别失败")
        messagebox.showerror("识别失败", f"识别过程出现错误：\n{err}")

    # ---------------- 表格提取 ----------------
    def start_table_extract(self):
        if self._recognizing:
            return
        if not self.image_path:
            messagebox.showinfo("提示", "请先选择一张包含表格的图片。")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("缺少依赖", "未安装 openpyxl，无法导出 Excel。\n请运行：pip install openpyxl")
            return
        self._recognizing = True
        self.select_btn.configure(state="disabled")
        self.recognize_btn.configure(state="disabled")
        self.table_btn.configure(state="disabled", text="提取中…")
        self.copy_btn.configure(state="disabled")
        self.result_box.delete("1.0", "end")
        self.char_count_label.configure(text="")
        self.progress_bar.start()
        self.set_status("正在识别表格并生成 Excel，请稍候…")
        threading.Thread(
            target=self._table_worker, args=(self.image_path,), daemon=True
        ).start()

    def _table_worker(self, path):
        try:
            items = recognize_with_boxes(path)
            table = extract_table(items)
            if not table:
                self._msg_queue.put(("table_empty", None))
                return
            # 生成桌面路径
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"表格提取结果_{timestamp}.xlsx"
            save_path = os.path.join(desktop, filename)
            export_table_to_excel(table, save_path)
            # 生成预览文本（制表符分隔）
            preview_lines = ["\t".join(row) for row in table]
            preview = "\n".join(preview_lines)
            self._msg_queue.put(("table_done", (save_path, preview, len(table), max(len(r) for r in table))))
        except Exception as exc:
            self._msg_queue.put(("table_error", str(exc)))

    def _on_table_done(self, payload):
        save_path, preview, rows, cols = payload
        self._recognizing = False
        self.select_btn.configure(state="normal")
        self.recognize_btn.configure(state="normal")
        self.table_btn.configure(state="normal", text="表格提取")
        self.progress_bar.stop()
        self.result_box.insert("1.0", preview)
        self.char_count_label.configure(text=f"{rows} 行 × {cols} 列")
        self.copy_btn.configure(state="normal")
        self.set_status(f"表格已保存到桌面：{os.path.basename(save_path)}")
        messagebox.showinfo("表格提取完成", f"已识别 {rows} 行 × {cols} 列的表格，\nExcel 文件已保存到桌面：\n\n{os.path.basename(save_path)}")

    def _on_table_empty(self):
        self._recognizing = False
        self.select_btn.configure(state="normal")
        self.recognize_btn.configure(state="normal")
        self.table_btn.configure(state="normal", text="表格提取")
        self.progress_bar.stop()
        self.set_status("未在图片中检测到表格内容")

    def _on_table_error(self, err):
        self._recognizing = False
        self.select_btn.configure(state="normal")
        self.recognize_btn.configure(state="normal")
        self.table_btn.configure(state="normal", text="表格提取")
        self.progress_bar.stop()
        self.set_status("表格提取失败")
        messagebox.showerror("表格提取失败", f"出现错误：\n{err}")

    def copy_text(self):
        text = self.result_box.get("1.0", "end").strip()
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.set_status("已复制到剪贴板，可直接粘贴使用")

    def clear_all(self):
        self.result_box.delete("1.0", "end")
        self.char_count_label.configure(text="")
        self.copy_btn.configure(state="disabled")
        self.set_status("已清空")

    def set_status(self, msg):
        self.status_label.configure(text=msg)


# ---------------------------------------------------------------------------
# 工具函数与入口
# ---------------------------------------------------------------------------
def resource_path(rel):
    """兼容源码运行与 PyInstaller 打包后的资源路径。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def run_selftest(argv):
    """隐藏自检模式：--selftest <图片路径> [输出文件路径]，供打包后验证使用。"""
    if not argv:
        return 2
    img_path, out_path = argv[0], (argv[1] if len(argv) > 1 else None)
    try:
        text = recognize_image(img_path)
    except Exception as exc:
        if out_path:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(f"ERROR: {exc}")
        return 1
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
    return 0


def main():
    if len(sys.argv) >= 2 and sys.argv[1] == "--selftest":
        sys.exit(run_selftest(sys.argv[2:]))

    # 高 DPI 适配，保证界面文字清晰
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    app = OCRApp()
    app.mainloop()


if __name__ == "__main__":
    main()
