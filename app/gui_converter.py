from __future__ import annotations

import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from markitdown import MarkItDown


class MarkItDownGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("文档转 Markdown")
        self.root.geometry("780x560")

        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=12, pady=12)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="输入文件").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.input_file, width=78).grid(
            row=1, column=0, columnspan=2, sticky="we", pady=6
        )
        tk.Button(frame, text="选择文件", command=self.choose_file, width=12).grid(
            row=1, column=2, padx=6
        )

        tk.Label(frame, text="输出目录").grid(row=2, column=0, sticky="w", pady=(12, 0))
        tk.Entry(frame, textvariable=self.output_dir, width=78).grid(
            row=3, column=0, columnspan=2, sticky="we", pady=6
        )
        tk.Button(frame, text="选择目录", command=self.choose_output_dir, width=12).grid(
            row=3, column=2, padx=6
        )

        tk.Button(
            frame,
            text="开始转换",
            command=self.start_convert,
            height=2
        ).grid(row=4, column=0, columnspan=3, sticky="we", pady=(16, 10))

        self.log_box = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=22)
        self.log_box.grid(row=5, column=0, columnspan=3, sticky="nsew")

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(5, weight=1)

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要转换的文件",
            filetypes=[
                ("支持的文件", "*.pdf *.docx *.pptx *.xlsx *.html *.htm *.csv *.json *.xml *.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.input_file.set(path)

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择导出目录")
        if path:
            self.output_dir.set(path)

    def log(self, message: str) -> None:
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)

    def convert_file(self) -> None:
        input_path = Path(self.input_file.get()).expanduser()
        output_dir = Path(self.output_dir.get()).expanduser()

        if not input_path.exists():
            self.log("错误：输入文件不存在")
            messagebox.showerror("错误", "输入文件不存在")
            return

        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.log(f"错误：无法创建输出目录：{e}")
                messagebox.showerror("错误", f"无法创建输出目录：{e}")
                return

        output_file = output_dir / f"{input_path.stem}.md"

        self.log(f"开始转换：{input_path}")
        self.log(f"输出文件：{output_file}")

        try:
            md = MarkItDown()
            result = md.convert(str(input_path))
            content = getattr(result, "text_content", "")

            if not content or not content.strip():
                self.log("转换失败：未提取到内容")
                messagebox.showwarning("提示", "未提取到内容，可能是扫描版 PDF 或内容格式不支持。")
                return

            output_file.write_text(content, encoding="utf-8")
            self.log("转换成功")
            messagebox.showinfo("完成", f"转换成功：\n{output_file}")
        except Exception as e:
            self.log(f"转换失败：{e}")
            messagebox.showerror("错误", f"转换失败：\n{e}")

    def start_convert(self) -> None:
        if not self.input_file.get().strip():
            messagebox.showwarning("提示", "请先选择输入文件")
            return

        if not self.output_dir.get().strip():
            messagebox.showwarning("提示", "请先选择输出目录")
            return

        thread = threading.Thread(target=self.convert_file, daemon=True)
        thread.start()


def main() -> None:
    root = tk.Tk()
    app = MarkItDownGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()