import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import datetime

# PDF設定
PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN_X = 50
MARGIN_Y = 50
LINE_HEIGHT = 30
WORD_SPACE = 2 * LINE_HEIGHT  # 空格2行高度
WORDS_PER_PAGE = 10
LINE_WIDTH = PAGE_WIDTH - 2 * MARGIN_X
DEFAULT_FONT = "Helvetica"
DEFAULT_FONT_SIZE = 14
FONT_OPTIONS = ["Helvetica", "Times-Roman", "Courier"]
FONT_SIZE_OPTIONS = [12, 14, 16, 18, 20, 24]

# 註冊思源黑體字型（確保有這個檔案）
if os.path.exists("NotoSansTC-Regular.ttf"):
    pdfmetrics.registerFont(TTFont("黑體", "NotoSansTC-Regular.ttf"))
    FONT_OPTIONS.append("黑體")


def generate_pdf(word_list, mode, output_path, font_name, font_size):
    c = canvas.Canvas(output_path, pagesize=A4)
    c.setFont(font_name, font_size)

    x = MARGIN_X
    y = PAGE_HEIGHT - MARGIN_Y

    count = 0
    page_number = 1
    today_str = datetime.datetime.today().strftime("%Y/%m/%d")

    for word_entry in word_list:
        if mode == "簡易版":
            word = word_entry
            text = word
        else:
            parts = word_entry.split("::")
            word = parts[0].strip()
            example = parts[1].strip() if len(parts) > 1 else "(請補上例句)"
            text = f"{word} - {example}"

        c.drawString(x, y, text)

        # 畫兩條虛線
        c.setStrokeColor(colors.grey)
        c.setDash(1, 2)
        line_start_x = x
        line_end_x = x + LINE_WIDTH * 0.9

        y_line1 = y - LINE_HEIGHT
        y_line2 = y - 2 * LINE_HEIGHT

        c.line(line_start_x, y_line1, line_end_x, y_line1)
        c.line(line_start_x, y_line2, line_end_x, y_line2)
        c.setDash()

        y -= (WORD_SPACE + LINE_HEIGHT)
        count += 1

        if count % WORDS_PER_PAGE == 0:
            # 加頁碼和製作日期
            c.setFont(font_name, 10)
            c.drawCentredString(PAGE_WIDTH / 2, MARGIN_Y / 2, f"第 {page_number} 頁")
            c.drawRightString(PAGE_WIDTH - MARGIN_X, MARGIN_Y / 2, f"製作日期：{today_str}")
            c.showPage()
            page_number += 1
            c.setFont(font_name, font_size)
            x = MARGIN_X
            y = PAGE_HEIGHT - MARGIN_Y

    # 最後一頁加頁碼和製作日期
    c.setFont(font_name, 10)
    c.drawCentredString(PAGE_WIDTH / 2, MARGIN_Y / 2, f"第 {page_number} 頁")
    c.drawRightString(PAGE_WIDTH - MARGIN_X, MARGIN_Y / 2, f"製作日期：{today_str}")

    c.save()


class WordPracticeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("英文單字練習簿產生器")

        # 模式選擇
        self.mode_var = tk.StringVar(value="簡易版")
        tk.Label(root, text="選擇模式:").pack()
        tk.Radiobutton(root, text="簡易版 (單字+空格)", variable=self.mode_var, value="簡易版").pack()
        tk.Radiobutton(root, text="進階版 (單字+例句+空格)", variable=self.mode_var, value="進階版").pack()

        # 字型選擇
        tk.Label(root, text="選擇字型:").pack()
        self.font_var = tk.StringVar(value=DEFAULT_FONT)
        self.font_menu = tk.OptionMenu(root, self.font_var, *FONT_OPTIONS)
        self.font_menu.pack()

        # 字型大小選擇
        tk.Label(root, text="選擇字型大小:").pack()
        self.font_size_var = tk.IntVar(value=DEFAULT_FONT_SIZE)
        self.font_size_menu = tk.OptionMenu(root, self.font_size_var, *FONT_SIZE_OPTIONS)
        self.font_size_menu.pack()

        # 單字輸入框
        tk.Label(root, text="請輸入單字（每行一個；進階版用 '單字::例句' 格式）:").pack()
        self.text_input = tk.Text(root, width=50, height=15)
        self.text_input.pack()

        # 產生按鈕
        tk.Button(root, text="產生PDF", command=self.create_pdf).pack(pady=10)

    def create_pdf(self):
        words_text = self.text_input.get("1.0", tk.END).strip()
        if not words_text:
            messagebox.showerror("錯誤", "請輸入至少一個單字！")
            return

        word_list = words_text.split("\n")
        mode = self.mode_var.get()

        font_name = self.font_var.get()
        font_size = self.font_size_var.get()

        if font_name == "黑體" and not os.path.exists("NotoSansTC-Regular.ttf"):
            messagebox.showerror("錯誤", "找不到黑體字型檔 NotoSansTC-Regular.ttf！請確認檔案存在於程式同一資料夾。")
            return

        # 自動產生檔名
        today = datetime.datetime.today().strftime("%Y%m%d")
        default_filename = f"單字練習簿_{today}.pdf"

        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_filename, filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            return

        generate_pdf(word_list, mode, file_path, font_name, font_size)
        messagebox.showinfo("成功", f"PDF已產生：\n{file_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = WordPracticeApp(root)
    root.mainloop()
