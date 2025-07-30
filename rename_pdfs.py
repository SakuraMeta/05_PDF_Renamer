import os
import shutil
import datetime
import configparser
import re
import tkinter as tk
from tkinter import ttk, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageTk

class PdfRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Renamer")
        self.root.geometry("500x450")

        self.load_config()
        self.setup_directories()

        self.pdf_files = sorted([f for f in os.listdir(self.input_dir) if f.lower().endswith('.pdf')])
        self.current_file_index = 0

        self.setup_ui()

        if not self.pdf_files:
            messagebox.showinfo("情報", f"{self.input_dir} にPDFファイルがありません。")
            self.root.quit()
        else:
            self.process_next_pdf()

    def load_config(self):
        config = configparser.ConfigParser()
        config.read('config.txt', encoding='utf-8')
        
        self.input_dir = config.get('Paths', 'input_dir', fallback='pdf_input')
        self.output_dir = config.get('Paths', 'output_dir', fallback='pdf_output')
        self.log_dir = config.get('Paths', 'log_dir', fallback='log_output')

        self.ocr_x = config.getint('OCR', 'x', fallback=50)
        self.ocr_y = config.getint('OCR', 'y', fallback=50)
        self.ocr_width = config.getint('OCR', 'width', fallback=200)
        self.ocr_height = config.getint('OCR', 'height', fallback=50)
        self.ocr_rect = fitz.Rect(self.ocr_x, self.ocr_y, self.ocr_x + self.ocr_width, self.ocr_y + self.ocr_height)
        
        self.filter_digits = config.getint('Filter', 'digits', fallback=0)

    def setup_directories(self):
        for dir_path in [self.input_dir, self.output_dir, self.log_dir]:
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

    def setup_ui(self):
        self.root.grid_columnconfigure(0, weight=1)

        # Status Label
        self.status_label = ttk.Label(self.root, text="", anchor="center")
        self.status_label.grid(row=0, column=0, pady=10, sticky="ew")

        # Image Label
        self.image_label = ttk.Label(self.root, text="PDFの範囲画像", relief="solid", anchor="center")
        self.image_label.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.root.grid_rowconfigure(1, weight=1)

        # Filename Entry
        ttk.Label(self.root, text="新しいファイル名 (拡張子不要):").grid(row=2, column=0, padx=20, pady=(10, 2), sticky="w")
        self.filename_var = tk.StringVar()
        self.filename_entry = ttk.Entry(self.root, textvariable=self.filename_var, width=60)
        self.filename_entry.grid(row=3, column=0, padx=20, pady=2, sticky="ew")

        # OK Button
        self.ok_button = ttk.Button(self.root, text="OK & 次へ ▶", command=self.on_ok_click, width=20)
        self.ok_button.grid(row=4, column=0, pady=20)

    def process_next_pdf(self):
        if self.current_file_index >= len(self.pdf_files):
            messagebox.showinfo("完了", "すべてのPDFファイルの処理が完了しました。")
            self.root.quit()
            return

        original_filename = self.pdf_files[self.current_file_index]
        self.status_label.config(text=f"処理中: {original_filename} ({self.current_file_index + 1}/{len(self.pdf_files)})")
        pdf_path = os.path.join(self.input_dir, original_filename)

        try:
            doc = fitz.open(pdf_path)
            page = doc.load_page(0)

            # 1. 画像の抽出と表示
            pix = page.get_pixmap(clip=self.ocr_rect, dpi=150)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # 画像をリサイズしてUIに合わせる
            img.thumbnail((400, 200), Image.Resampling.LANCZOS)
            self.photo_image = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.photo_image)

            # 2. テキストの抽出とフィルタリング
            text = page.get_text(clip=self.ocr_rect, sort=True).strip()
            numbers = re.findall(r'\d+', text)
            extracted_text = "".join(numbers)

            if self.filter_digits > 0 and len(extracted_text) != self.filter_digits:
                self.filename_var.set(f"(桁数エラー: {extracted_text})")
            else:
                self.filename_var.set(extracted_text)
            
            doc.close()
            self.filename_entry.focus()
            self.filename_entry.select_range(0, 'end')

        except Exception as e:
            messagebox.showerror("エラー", f"{original_filename}の処理中にエラーが発生しました:\n{e}")
            self.current_file_index += 1
            self.process_next_pdf()

    def on_ok_click(self):
        new_filename_base = self.filename_var.get().strip()
        if not new_filename_base:
            messagebox.showwarning("警告", "ファイル名を入力してください。")
            return

        original_filename = self.pdf_files[self.current_file_index]
        original_path = os.path.join(self.input_dir, original_filename)
        new_filename = f"{new_filename_base}.pdf"
        new_path = os.path.join(self.output_dir, new_filename)

        if os.path.exists(new_path):
            if not messagebox.askyesno("確認", f"{new_filename} は既に存在します。上書きしますか？"):
                return

        try:
            # ファイルをコピーしてリネーム
            shutil.copy2(original_path, new_path)
            # ログを記録
            self.write_log(new_filename_base)
            
            self.current_file_index += 1
            self.process_next_pdf()

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました:\n{e}")

    def write_log(self, text_to_log):
        log_filename = os.path.join(self.log_dir, f"{datetime.date.today().strftime('%Y%m%d')}.txt")
        with open(log_filename, 'a', encoding='utf-8') as f:
            f.write(f"{text_to_log}\n")

if __name__ == '__main__':
    root = tk.Tk()
    app = PdfRenamerApp(root)
    root.mainloop()
