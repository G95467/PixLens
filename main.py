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
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageOps

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
        header.grid_columnconfigure(2, weight=0)

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
        btn_row.grid_columnconfigure(1, weight=2)
        btn_row.grid_columnconfigure(2, weight=2)

        self.recognize_btn = ctk.CTkButton(
            btn_row, text="开始识别", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            corner_radius=10, command=self.start_recognize,
        )
        self.recognize_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.copy_btn = ctk.CTkButton(
            btn_row, text="复制文字", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            corner_radius=10, fg_color="#2E7D32", hover_color="#388E3C",
            state="disabled", command=self.copy_text,
        )
        self.copy_btn.grid(row=0, column=1, sticky="ew", padx=6)

        self.clear_btn = ctk.CTkButton(
            btn_row, text="清空", height=42,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            corner_radius=10, fg_color="#3A3F4B", hover_color="#4A5060",
            command=self.clear_all,
        )
        self.clear_btn.grid(row=0, column=2, sticky="ew", padx=(6, 0))

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
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _on_done(self, text):
        self._recognizing = False
        self.select_btn.configure(state="normal")
        self.recognize_btn.configure(state="normal", text="开始识别")
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
        self.progress_bar.stop()
        self.set_status("识别失败")
        messagebox.showerror("识别失败", f"识别过程出现错误：\n{err}")

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
