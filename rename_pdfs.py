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
        self.root.state('zoomed') # ウィンドウを最大化

        # --- Style ---
        style = ttk.Style(self.root)
        try:
            style.theme_use('vista') # Use a modern theme if available
        except tk.TclError:
            print("Vista theme not available, using default.")

        # Configure a standard button style
        style.configure('TButton', font=('Yu Gothic UI', 9), padding=5)

        self.load_config()
        self.setup_directories()

        self.pdf_files = sorted([f for f in os.listdir(self.input_dir) if f.lower().endswith('.pdf')])
        self.current_file_index = 0
        self.selection_mode = False
        self.rect_start_x = None
        self.rect_start_y = None
        self.current_rect_id = None
        self.image_scale_factor = 1.0
        self.initial_load_done = False # 初回ロード管理フラグ

        self.setup_ui()

        if not self.pdf_files:
            messagebox.showinfo("情報", f"{self.input_dir} にPDFファイルがありません。")
            self.root.quit()
        # 初回のPDF読み込みは on_canvas_configure で行う

    def load_config(self):
        self.config = configparser.ConfigParser()
        self.config.read('config.txt', encoding='utf-8')
        
        self.input_dir = self.config.get('Paths', 'input_dir', fallback='pdf_input')
        self.output_dir = self.config.get('Paths', 'output_dir', fallback='pdf_output')
        self.log_dir = self.config.get('Paths', 'log_dir', fallback='log_output')

        self.ocr_x = self.config.getint('OCR', 'x', fallback=50)
        self.ocr_y = self.config.getint('OCR', 'y', fallback=50)
        self.ocr_width = self.config.getint('OCR', 'width', fallback=200)
        self.ocr_height = self.config.getint('OCR', 'height', fallback=50)
        self.ocr_rect = fitz.Rect(self.ocr_x, self.ocr_y, self.ocr_x + self.ocr_width, self.ocr_y + self.ocr_height)
        
        self.filter_digits = self.config.getint('Filter', 'digits', fallback=0)

    def setup_directories(self):
        for dir_path in [self.input_dir, self.output_dir, self.log_dir]:
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

    def setup_ui(self):
        # --- Layout using Pack --- 
        # Top Frame for status and button
        top_frame = ttk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))

        self.status_label = ttk.Label(top_frame, text="", anchor="center")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.reset_button = ttk.Button(top_frame, text="読み取り範囲の再設定", command=self.toggle_selection_mode)
        self.reset_button.pack(side=tk.RIGHT)

        # Bottom Frame for controls
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5, 10))

        # Config button on the far right, anchored to the bottom of its space
        config_button = ttk.Button(bottom_frame, text="config.txtを開く", command=self.open_config_file)
        config_button.pack(side=tk.RIGHT, anchor='s', padx=5, pady=5)

        # Central controls frame, will be centered in the remaining space
        controls_frame = ttk.Frame(bottom_frame)
        controls_frame.pack(expand=True) # expand=True helps with centering

        ttk.Label(controls_frame, text="新しいファイル名 (拡張子不要):").pack() # Label on top

        # Frame for entry and button to be side-by-side
        input_line_frame = ttk.Frame(controls_frame)
        input_line_frame.pack(pady=5)

        self.filename_var = tk.StringVar()
        self.filename_entry = ttk.Entry(input_line_frame, textvariable=self.filename_var, width=60, justify='center')
        self.filename_entry.pack(side=tk.LEFT, padx=(0, 5))

        self.ok_button = ttk.Button(input_line_frame, text="OK & 次へ ▶", command=self.on_ok_click, width=20)
        self.ok_button.pack(side=tk.LEFT)

        # Canvas for Image (fills remaining space)
        canvas_frame = ttk.Frame(self.root, relief="sunken", borderwidth=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.canvas = tk.Canvas(canvas_frame, background="#f0f0f0")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

    def open_config_file(self):
        config_path = os.path.abspath('config.txt')
        try:
            os.startfile(config_path)
        except FileNotFoundError:
            messagebox.showerror("エラー", f"config.txtが見つかりません。\nパス: {config_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルを開けませんでした。\n{e}")

    def process_next_pdf(self):
        if self.current_file_index >= len(self.pdf_files):
            messagebox.showinfo("完了", "すべてのPDFファイルの処理が完了しました。")
            self.root.quit()
            return

        self.selection_mode = False # Reset selection mode for each new PDF
        # self.reset_button.config(relief="raised") # ttk.Buttonではreliefはサポートされていないため削除

        original_filename = self.pdf_files[self.current_file_index]
        self.status_label.config(text=f"処理中: {original_filename} ({self.current_file_index + 1}/{len(self.pdf_files)})")
        self.pdf_path = os.path.join(self.input_dir, original_filename)

        try:
            self.doc = fitz.open(self.pdf_path)
            page = self.doc.load_page(0)
            self.page_rect = page.rect

            # Display full page image
            self.display_full_page(page)
            # Extract text from the current rect
            self.extract_text_from_rect()

            self.filename_entry.focus()
            self.filename_entry.select_range(0, 'end')

        except Exception as e:
            messagebox.showerror("エラー", f"{original_filename}の処理中にエラーが発生しました:\n{e}")
            self.current_file_index += 1
            self.process_next_pdf()
    
    def display_full_page(self, page):
        # Render the full page
        canvas_w, canvas_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if canvas_w < 50 or canvas_h < 50: # Canvas not ready
            canvas_w, canvas_h = 800, 1000 # Default size
        
        zoom_x = canvas_w / self.page_rect.width
        zoom_y = canvas_h / self.page_rect.height
        self.image_scale_factor = min(zoom_x, zoom_y) * 0.95 # Add some padding

        mat = fitz.Matrix(self.image_scale_factor, self.image_scale_factor)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        self.canvas.delete("all")
        self.photo_image = ImageTk.PhotoImage(img)
        self.canvas.create_image(canvas_w/2, canvas_h/2, image=self.photo_image, anchor='center')
        self.draw_ocr_rect()

    def draw_ocr_rect(self):
        if self.current_rect_id:
            self.canvas.delete(self.current_rect_id)
        
        canvas_w, canvas_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        img_w, img_h = self.photo_image.width(), self.photo_image.height()
        offset_x = (canvas_w - img_w) / 2
        offset_y = (canvas_h - img_h) / 2

        x0 = self.ocr_rect.x0 * self.image_scale_factor + offset_x
        y0 = self.ocr_rect.y0 * self.image_scale_factor + offset_y
        x1 = self.ocr_rect.x1 * self.image_scale_factor + offset_x
        y1 = self.ocr_rect.y1 * self.image_scale_factor + offset_y
        self.current_rect_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", width=2)

    def extract_text_from_rect(self):
        page = self.doc.load_page(0)
        text = page.get_text(clip=self.ocr_rect, sort=True).strip()
        numbers = re.findall(r'\d+', text)
        extracted_text = "".join(numbers)

        if self.filter_digits > 0 and len(extracted_text) != self.filter_digits:
            self.filename_var.set(f"(桁数エラー: {extracted_text})")
        else:
            self.filename_var.set(extracted_text)

    def on_ok_click(self):
        if self.selection_mode:
            self.toggle_selection_mode() # Exit selection mode if active
            return
        
        new_filename_base = self.filename_var.get().strip()
        if not new_filename_base or new_filename_base.startswith("("):
            messagebox.showwarning("警告", "有効なファイル名を入力してください。")
            return

        new_filename = f"{new_filename_base}.pdf"
        new_path = os.path.join(self.output_dir, new_filename)

        if os.path.exists(new_path):
            if not messagebox.askyesno("確認", f"{new_filename} は既に存在します。上書きしますか？"):
                return

        try:
            self.doc.close() # Close the document before copying
            shutil.copy2(self.pdf_path, new_path)
            self.write_log(new_filename_base)
            self.current_file_index += 1
            self.process_next_pdf()
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました:\n{e}")

    def write_log(self, text_to_log):
        log_filename = os.path.join(self.log_dir, f"{datetime.date.today().strftime('%Y%m%d')}.txt")
        with open(log_filename, 'a', encoding='utf-8') as f:
            f.write(f"{text_to_log}\n")

    def toggle_selection_mode(self):
        self.selection_mode = not self.selection_mode
        if self.selection_mode:
            # self.reset_button.config(relief="sunken") # ttk.Buttonではreliefはサポートされていないため削除
            self.status_label.config(text="範囲を選択してください: マウスをドラッグして赤い四角を作成します。")
            self.canvas.config(cursor="crosshair")
        else:
            # self.reset_button.config(relief="raised") # ttk.Buttonではreliefはサポートされていないため削除
            original_filename = self.pdf_files[self.current_file_index]
            self.status_label.config(text=f"処理中: {original_filename} ({self.current_file_index + 1}/{len(self.pdf_files)})")
            self.canvas.config(cursor="")

    def on_mouse_down(self, event):
        if not self.selection_mode: return
        self.rect_start_x = event.x
        self.rect_start_y = event.y
        if self.current_rect_id:
            self.canvas.delete(self.current_rect_id)
        self.current_rect_id = self.canvas.create_rectangle(self.rect_start_x, self.rect_start_y, self.rect_start_x, self.rect_start_y, outline="red", width=2)

    def on_mouse_drag(self, event):
        if not self.selection_mode or self.rect_start_x is None: return
        self.canvas.coords(self.current_rect_id, self.rect_start_x, self.rect_start_y, event.x, event.y)

    def on_mouse_up(self, event):
        if not self.selection_mode or self.rect_start_x is None: return
        canvas_w, canvas_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        img_w, img_h = self.photo_image.width(), self.photo_image.height()
        offset_x = (canvas_w - img_w) / 2
        offset_y = (canvas_h - img_h) / 2

        # Convert canvas coords to PDF coords
        x0 = (min(self.rect_start_x, event.x) - offset_x) / self.image_scale_factor
        y0 = (min(self.rect_start_y, event.y) - offset_y) / self.image_scale_factor
        x1 = (max(self.rect_start_x, event.x) - offset_x) / self.image_scale_factor
        y1 = (max(self.rect_start_y, event.y) - offset_y) / self.image_scale_factor
        
        # Clamp to page boundaries
        x0 = max(0, x0)
        y0 = max(0, y0)
        x1 = min(self.page_rect.width, x1)
        y1 = min(self.page_rect.height, y1)

        # 1. 新しいOCR範囲をインスタンス変数に設定
        self.ocr_rect = fitz.Rect(x0, y0, x1, y1)

        # 2. 新しい範囲からテキストを抽出し、ファイル名ボックスを更新
        self.extract_text_from_rect()

        # 3. 新しい範囲の座標をconfig.txtに保存
        self.save_new_config(x0, y0, x1 - x0, y1 - y0)

        # 4. 選択モードを終了
        self.toggle_selection_mode()
        self.rect_start_x, self.rect_start_y = None, None

    def save_new_config(self, x, y, w, h):
        self.config.set('OCR', 'x', str(int(x)))
        self.config.set('OCR', 'y', str(int(y)))
        self.config.set('OCR', 'width', str(int(w)))
        self.config.set('OCR', 'height', str(int(h)))
        with open('config.txt', 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)
        messagebox.showinfo("成功", "新しい読み取り範囲をconfig.txtに保存しました。")

    def on_canvas_configure(self, event):
        if not self.initial_load_done and self.pdf_files:
            self.initial_load_done = True
            self.process_next_pdf()

if __name__ == '__main__':
    root = tk.Tk()
    app = PdfRenamerApp(root)
    root.mainloop()

