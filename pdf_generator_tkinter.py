# ✅ 整合 ims_list.json 的商品明細查詢 + PDF 生成（包含批次與單筆明細）
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os, json, platform, subprocess

# 註冊中文字體避免黑框（Windows: 微軟正黑體）
try:
    pdfmetrics.registerFont(TTFont('CustomFont', 'msjhbd.ttc'))  # 使用粗體版本
    FONT = 'CustomFont'
except:
    FONT = 'Helvetica-Bold'

# 讀取 ims_list.json
with open("ims_list.json", "r", encoding="utf-8") as f:
    ims_data = json.load(f)
    ims_lookup = {item["Item No"].strip(): item["Item Description"].strip() for item in ims_data}

class PDFGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF文件生成器（含商品明細）")
        self.root.geometry("900x700")
        self.items = []

        self.save_path = tk.StringVar(value=os.getcwd())
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.sender_store_var = tk.StringVar()
        self.sender_name_var = tk.StringVar()
        self.receiver_store_var = tk.StringVar()
        self.receiver_name_var = tk.StringVar()

        self.item_code_var = tk.StringVar()
        self.item_desc_var = tk.StringVar()
        self.item_qty_var = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # 基本資訊
        info = [
            ("日期", self.date_var),
            ("寄出店別", self.sender_store_var),
            ("寄件人", self.sender_name_var),
            ("收件店別", self.receiver_store_var),
            ("收件人", self.receiver_name_var),
        ]
        for i, (label, var) in enumerate(info):
            ttk.Label(frame, text=label, width=12, anchor="e").grid(row=i, column=0, sticky="e")
            ttk.Entry(frame, textvariable=var, width=50).grid(row=i, column=1, pady=3, sticky="w")

        # 商品輸入
        ttk.Label(frame, text="貨號", width=12, anchor="e").grid(row=5, column=0, sticky="e")
        code_entry = ttk.Entry(frame, textvariable=self.item_code_var, width=20)
        code_entry.grid(row=5, column=1, sticky="w")
        code_entry.bind("<FocusOut>", self.fill_description)

        ttk.Label(frame, text="品名", width=12, anchor="e").grid(row=6, column=0, sticky="e")
        ttk.Entry(frame, textvariable=self.item_desc_var, state="readonly", width=70).grid(row=6, column=1, sticky="w")

        ttk.Label(frame, text="數量", width=12, anchor="e").grid(row=7, column=0, sticky="e")
        ttk.Entry(frame, textvariable=self.item_qty_var, width=10).grid(row=7, column=1, sticky="w")

        ttk.Button(frame, text="加入商品", command=self.add_item).grid(row=8, column=1, sticky="e", pady=5)

        # 商品列表
        self.tree = ttk.Treeview(frame, columns=("code", "desc", "qty"), show="headings", height=7)
        self.tree.heading("code", text="貨號")
        self.tree.heading("desc", text="品名")
        self.tree.heading("qty", text="數量")
        self.tree.column("code", width=120)
        self.tree.column("desc", width=580)
        self.tree.column("qty", width=80)
        self.tree.grid(row=9, column=0, columnspan=2, pady=10)

        # 按鈕
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=10, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="選擇儲存資料夾", command=self.choose_path).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="產生PDF", command=self.generate_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="清除所有商品", command=self.clear_items).pack(side=tk.LEFT, padx=5)

    def choose_path(self):
        path = filedialog.askdirectory()
        if path:
            self.save_path.set(path)

    def fill_description(self, event=None):
        code = self.item_code_var.get().strip()
        self.item_desc_var.set(ims_lookup.get(code, ""))

    def add_item(self):
        code = self.item_code_var.get().strip()
        desc = self.item_desc_var.get().strip()
        qty = self.item_qty_var.get().strip()
        if not code or not desc or not qty:
            messagebox.showerror("錯誤", "請完整填寫商品欄位")
            return
        self.items.append({"code": code, "desc": desc, "qty": qty})
        self.tree.insert("", "end", values=(code, desc, qty))
        self.item_code_var.set("")
        self.item_desc_var.set("")
        self.item_qty_var.set("")

    def clear_items(self):
        self.items.clear()
        for i in self.tree.get_children():
            self.tree.delete(i)

    def generate_pdf(self):
        if not self.items:
            messagebox.showwarning("警告", "請加入至少一項商品")
            return

        filename = f"transfer_{self.date_var.get().replace('/', '-')}_{self.sender_store_var.get()}_to_{self.receiver_store_var.get()}.pdf"
        filepath = os.path.join(self.save_path.get(), filename)

        c = canvas.Canvas(filepath, pagesize=landscape(A4))
        width, height = landscape(A4)

        c.setFont(FONT, 22)
        c.drawString(100, height - 50, "Transfer Document / 調貨單")

        c.setFont(FONT, 16)
        y = height - 100
        for label, value in [
            ("Date", self.date_var.get()),
            ("From Store", self.sender_store_var.get()),
            ("Sender", self.sender_name_var.get()),
            ("To Store", self.receiver_store_var.get()),
            ("Receiver", self.receiver_name_var.get())
        ]:
            c.drawString(60, y, f"{label}: {value}")
            y -= 30

        # 商品明細
        c.setFont(FONT, 16)
        c.drawString(60, y, "Items:")
        y -= 25
        c.setFont(FONT, 14)
        c.drawString(60, y, "貨號")
        c.drawString(200, y, "品名")
        c.drawString(600, y, "數量")
        y -= 20
        for item in self.items:
            c.drawString(60, y, item['code'])
            c.drawString(200, y, item['desc'][:60])
            c.drawString(600, y, str(item['qty']))
            y -= 20
            if y < 60:
                c.showPage()
                y = height - 80

        c.save()
        messagebox.showinfo("完成", f"PDF 已儲存至：\n{filepath}")

        if messagebox.askyesno("開啟資料夾", "是否開啟儲存資料夾？"):
            self.open_folder(self.save_path.get())

    def open_folder(self, path):
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFGeneratorApp(root)
    root.mainloop()
