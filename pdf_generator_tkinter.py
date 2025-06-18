# ✅ 整合 ims_list.json 的商品明細查詢 + PDF 生成（包含批次與單筆明細）
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os
import json
import platform
import subprocess


class PDFGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF調貨單生成器 - Transfer Document Generator")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)
        self.root.minsize(900, 700)

        # 初始化變數
        self.items = []
        self.font_loaded = False

        # 設置UI變數
        self.setup_variables()

        # 設置字體
        self.setup_fonts()

        # 載入IMS數據
        self.load_ims_data()

        # 設置UI
        self.setup_ui()

        # 綁定事件
        self.bind_events()

    def setup_variables(self):
        """初始化UI變數"""
        self.save_path = tk.StringVar(value=os.getcwd())
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.sender_store_var = tk.StringVar()
        self.sender_name_var = tk.StringVar()
        self.receiver_store_var = tk.StringVar()
        self.receiver_name_var = tk.StringVar()
        self.item_code_var = tk.StringVar()
        self.item_desc_var = tk.StringVar()
        self.item_qty_var = tk.StringVar()
        self.notes_var = tk.StringVar()

    def setup_fonts(self):
        """設置中文字體 - 跨平台支援"""
        try:
            system = platform.system()
            font_paths = []

            if system == "Windows":
                font_paths = [
                    "C:/Windows/Fonts/msjh.ttc",    # 微軟正黑體
                    "C:/Windows/Fonts/msyh.ttc",    # 微軟雅黑
                    "C:/Windows/Fonts/simhei.ttf",  # 黑體
                    "C:/Windows/Fonts/simsun.ttc",  # 宋體
                ]
            elif system == "Darwin":  # macOS
                font_paths = [
                    "/System/Library/Fonts/PingFang.ttc",
                    "/System/Library/Fonts/Helvetica.ttc",
                    "/System/Library/Fonts/Supplemental/Songti.ttc",
                ]
            else:  # Linux
                font_paths = [
                    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                ]

            # 嘗試載入字體
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(
                            TTFont('ChineseFont', font_path))
                        self.font_name = 'ChineseFont'
                        self.font_loaded = True
                        print(f"成功載入字體: {font_path}")
                        break
                    except Exception as e:
                        print(f"字體載入失敗 {font_path}: {e}")
                        continue

            if not self.font_loaded:
                print("未找到合適的中文字體，使用預設字體")
                self.font_name = 'Helvetica'

        except Exception as e:
            print(f"字體設置錯誤: {e}")
            self.font_name = 'Helvetica'

    def load_ims_data(self):
        """載入IMS數據"""
        self.ims_lookup = {}
        try:
            # 嘗試不同路徑
            json_paths = [
                "ims_list.json",
                os.path.join(os.path.dirname(__file__), "ims_list.json"),
                os.path.join(os.getcwd(), "ims_list.json")
            ]

            for json_path in json_paths:
                if os.path.exists(json_path):
                    with open(json_path, "r", encoding="utf-8") as f:
                        ims_data = json.load(f)
                        self.ims_lookup = {
                            item["Item No"].strip(): item["Item Description"].strip()
                            for item in ims_data if "Item No" in item and "Item Description" in item
                        }
                    print(f"成功載入 {len(self.ims_lookup)} 筆商品資料")
                    return

            print("警告: 未找到 ims_list.json 檔案")
            messagebox.showwarning(
                "警告", "未找到商品資料檔案 (ims_list.json)\n商品查詢功能將無法使用")

        except Exception as e:
            print(f"載入商品資料錯誤: {e}")
            messagebox.showerror("錯誤", f"載入商品資料失敗: {e}")

    def setup_ui(self):
        """設置使用者介面"""
        # 清除舊內容
        for widget in self.root.winfo_children():
            widget.destroy()

        # 主標題
        self.create_header()

        # 主要內容
        self.create_main_content()

        # 按鈕區域
        self.create_buttons()

        # 狀態列
        self.create_status_bar()

    def create_header(self):
        """創建標題區域"""
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=5)

        title_label = ttk.Label(header_frame, text="調貨單 PDF 生成器",
                                font=('Arial', 18, 'bold'))
        title_label.pack()

        subtitle_label = ttk.Label(header_frame, text="Transfer Document Generator",
                                   font=('Arial', 10))
        subtitle_label.pack()

    def create_main_content(self):
        """創建主要內容區域"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 左側：基本資訊
        left_frame = ttk.LabelFrame(main_frame, text="基本資訊", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.create_basic_info(left_frame)

        # 右側：商品資訊
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.create_item_section(right_frame)

    def create_basic_info(self, parent):
        """創建基本資訊區域"""
        info_fields = [
            ("日期 (Date)", self.date_var),
            ("寄出店別 (From Store)", self.sender_store_var),
            ("寄件人 (Sender)", self.sender_name_var),
            ("收件店別 (To Store)", self.receiver_store_var),
            ("收件人 (Receiver)", self.receiver_name_var),
            ("備註 (Notes)", self.notes_var),
        ]

        for i, (label, var) in enumerate(info_fields):
            ttk.Label(parent, text=label).grid(
                row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(parent, textvariable=var, width=25)
            entry.grid(row=i, column=1, sticky="ew", pady=2, padx=(5, 0))

        parent.columnconfigure(1, weight=1)

        # 儲存路徑
        ttk.Label(parent, text="儲存路徑 (Save Path)").grid(
            row=len(info_fields), column=0, sticky="w", pady=2)
        path_frame = ttk.Frame(parent)
        path_frame.grid(row=len(info_fields), column=1,
                        sticky="ew", pady=2, padx=(5, 0))

        self.path_label = ttk.Label(path_frame, text=self.save_path.get(),
                                    relief="sunken", width=20)
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(path_frame, text="選擇", command=self.choose_path,
                   width=8).pack(side=tk.RIGHT, padx=(5, 0))

    def create_item_section(self, parent):
        """創建商品區域"""
        # 商品輸入區域
        item_input_frame = ttk.LabelFrame(parent, text="商品輸入", padding=10)
        item_input_frame.pack(fill=tk.X, pady=(0, 5))

        # 商品編號
        ttk.Label(item_input_frame, text="商品編號 (Item Code)").grid(
            row=0, column=0, sticky="w")
        self.code_entry = ttk.Entry(
            item_input_frame, textvariable=self.item_code_var, width=20)
        self.code_entry.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        ttk.Button(item_input_frame, text="查詢", command=self.lookup_item,
                   width=8).grid(row=0, column=2, padx=(5, 0))

        # 商品描述
        ttk.Label(item_input_frame, text="商品描述 (Description)").grid(
            row=1, column=0, sticky="w", pady=(5, 0))
        self.desc_entry = ttk.Entry(item_input_frame, textvariable=self.item_desc_var,
                                    state="readonly", width=40)
        self.desc_entry.grid(row=1, column=1, columnspan=2,
                             sticky="ew", pady=(5, 0), padx=(5, 0))

        # 數量
        ttk.Label(item_input_frame, text="數量 (Quantity)").grid(
            row=2, column=0, sticky="w", pady=(5, 0))
        self.qty_entry = ttk.Entry(
            item_input_frame, textvariable=self.item_qty_var, width=10)
        self.qty_entry.grid(row=2, column=1, sticky="w",
                            pady=(5, 0), padx=(5, 0))
        ttk.Button(item_input_frame, text="加入", command=self.add_item,
                   width=8).grid(row=2, column=2, padx=(5, 0), pady=(5, 0))

        item_input_frame.columnconfigure(1, weight=1)

        # 商品列表區域
        list_frame = ttk.LabelFrame(parent, text="商品清單", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 商品樹狀檢視
        columns = ("code", "desc", "qty")
        self.tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", height=10)

        self.tree.heading("code", text="商品編號")
        self.tree.heading("desc", text="商品描述")
        self.tree.heading("qty", text="數量")

        self.tree.column("code", width=120, minwidth=80)
        self.tree.column("desc", width=400, minwidth=200)
        self.tree.column("qty", width=80, minwidth=60)

        # 滾動條
        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 商品列表按鈕
        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(list_btn_frame, text="移除選取", command=self.remove_item).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(list_btn_frame, text="清空全部",
                   command=self.clear_items).pack(side=tk.LEFT)

    def create_buttons(self):
        """創建按鈕區域"""
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(btn_frame, text="清除表單", command=self.clear_form).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="產生 PDF", command=self.generate_pdf).pack(
            side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="預覽資料", command=self.preview_data).pack(
            side=tk.RIGHT, padx=(5, 0))

    def create_status_bar(self):
        """創建狀態列"""
        self.status_bar = ttk.Label(
            self.root, text="就緒", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def bind_events(self):
        """綁定事件"""
        self.code_entry.bind("<Return>", lambda e: self.lookup_item())
        self.code_entry.bind("<FocusOut>", lambda e: self.lookup_item())
        self.qty_entry.bind("<Return>", lambda e: self.add_item())

        # 雙擊編輯
        self.tree.bind("<Double-1>", self.edit_item)

    def choose_path(self):
        """選擇儲存路徑"""
        path = filedialog.askdirectory(title="選擇儲存資料夾")
        if path:
            self.save_path.set(path)
            self.path_label.config(text=path)
            self.status_bar.config(text=f"儲存路徑已設定: {path}")

    def lookup_item(self):
        """查詢商品資訊"""
        code = self.item_code_var.get().strip()
        if not code:
            self.item_desc_var.set("")
            return

        if code in self.ims_lookup:
            description = self.ims_lookup[code]
            self.item_desc_var.set(description)
            self.status_bar.config(text=f"找到商品: {code}")
        else:
            self.item_desc_var.set("未找到商品資訊")
            self.status_bar.config(text=f"未找到商品: {code}")

    def add_item(self):
        """加入商品到清單"""
        code = self.item_code_var.get().strip()
        desc = self.item_desc_var.get().strip()
        qty = self.item_qty_var.get().strip()

        if not code:
            messagebox.showerror("錯誤", "請輸入商品編號")
            return
        if not qty:
            messagebox.showerror("錯誤", "請輸入數量")
            return

        try:
            int(qty)  # 驗證數量是否為數字
        except ValueError:
            messagebox.showerror("錯誤", "數量必須是數字")
            return

        # 檢查是否已存在
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][0] == code:
                if messagebox.askyesno("確認", f"商品 {code} 已存在，是否要更新數量？"):
                    self.tree.item(item, values=(code, desc, qty))
                    self.clear_item_inputs()
                    self.status_bar.config(text=f"已更新商品: {code}")
                    return
                else:
                    return

        # 加入新商品
        self.tree.insert("", "end", values=(code, desc, qty))
        self.clear_item_inputs()
        self.status_bar.config(text=f"已加入商品: {code}")

    def clear_item_inputs(self):
        """清空商品輸入欄位"""
        self.item_code_var.set("")
        self.item_desc_var.set("")
        self.item_qty_var.set("")

    def remove_item(self):
        """移除選取的商品"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要移除的商品")
            return

        for item in selected:
            self.tree.delete(item)

        self.status_bar.config(text="已移除選取的商品")

    def clear_items(self):
        """清空所有商品"""
        if not self.tree.get_children():
            return

        if messagebox.askyesno("確認", "確定要清空所有商品嗎？"):
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.status_bar.config(text="已清空所有商品")

    def edit_item(self, event):
        """編輯商品項目"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return

        values = self.tree.item(item)['values']
        self.item_code_var.set(values[0])
        self.item_desc_var.set(values[1])
        self.item_qty_var.set(values[2])

        # 移除原項目
        self.tree.delete(item)

    def clear_form(self):
        """清空表單"""
        if messagebox.askyesno("確認", "確定要清空所有資料嗎？"):
            self.sender_store_var.set("")
            self.sender_name_var.set("")
            self.receiver_store_var.set("")
            self.receiver_name_var.set("")
            self.notes_var.set("")
            self.clear_items()
            self.date_var.set(datetime.now().strftime("%Y/%m/%d"))
            self.status_bar.config(text="表單已清空")

    def preview_data(self):
        """預覽資料"""
        if not self.validate_inputs():
            return

        items = [(self.tree.item(item)['values'][0],
                 self.tree.item(item)['values'][1],
                 self.tree.item(item)['values'][2])
                 for item in self.tree.get_children()]

        preview_text = f"""
調貨單預覽
================
日期: {self.date_var.get()}
寄出店別: {self.sender_store_var.get()}
寄件人: {self.sender_name_var.get()}
收件店別: {self.receiver_store_var.get()}
收件人: {self.receiver_name_var.get()}
備註: {self.notes_var.get()}

商品明細:
{'編號':<15} {'描述':<50} {'數量':<10}
{'-'*75}
"""

        for code, desc, qty in items:
            preview_text += f"{code:<15} {desc[:50]:<50} {qty:<10}\n"

        # 顯示預覽視窗
        preview_window = tk.Toplevel(self.root)
        preview_window.title("資料預覽")
        preview_window.geometry("800x600")

        text_widget = tk.Text(
            preview_window, wrap=tk.WORD, font=('Courier', 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar_preview = ttk.Scrollbar(
            preview_window, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar_preview.set)
        scrollbar_preview.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget.insert(tk.END, preview_text)
        text_widget.config(state=tk.DISABLED)

    def validate_inputs(self):
        """驗證輸入"""
        if not self.sender_store_var.get().strip():
            messagebox.showerror("錯誤", "請輸入寄出店別")
            return False
        if not self.receiver_store_var.get().strip():
            messagebox.showerror("錯誤", "請輸入收件店別")
            return False
        if not self.tree.get_children():
            messagebox.showerror("錯誤", "請至少加入一項商品")
            return False
        return True

    def generate_pdf(self):
        """生成PDF文件"""
        if not self.validate_inputs():
            return

        try:
            # 生成檔名
            date_str = self.date_var.get().replace('/', '-')
            filename = f"調貨單_{date_str}_{self.sender_store_var.get()}_to_{self.receiver_store_var.get()}.pdf"
            filepath = os.path.join(self.save_path.get(), filename)

            # 創建PDF
            self.create_pdf(filepath)

            self.status_bar.config(text=f"PDF已生成: {filename}")

            # 詢問是否開啟
            if messagebox.askyesno("完成", f"PDF已成功生成！\n\n檔案位置: {filepath}\n\n是否要開啟檔案？"):
                self.open_file(filepath)

            if messagebox.askyesno("開啟資料夾", "是否要開啟儲存資料夾？"):
                self.open_folder(self.save_path.get())

        except Exception as e:
            messagebox.showerror("錯誤", f"PDF生成失敗: {str(e)}")
            self.status_bar.config(text="PDF生成失敗")

    def create_pdf(self, filepath):
        """創建PDF文件"""
        c = canvas.Canvas(filepath, pagesize=landscape(A4))
        width, height = landscape(A4)

        # 設置字體
        font_size_title = 22
        font_size_header = 16
        font_size_content = 14
        font_size_small = 12

        # 標題
        c.setFont(self.font_name, font_size_title)
        title_text = "Transfer Document / 調貨單"
        title_width = c.stringWidth(
            title_text, self.font_name, font_size_title)
        c.drawString((width - title_width) / 2, height - 50, title_text)

        # 基本資訊
        c.setFont(self.font_name, font_size_header)
        y_pos = height - 100
        line_height = 25

        info_items = [
            ("Date / 日期", self.date_var.get()),
            ("From Store / 寄出店別", self.sender_store_var.get()),
            ("Sender / 寄件人", self.sender_name_var.get()),
            ("To Store / 收件店別", self.receiver_store_var.get()),
            ("Receiver / 收件人", self.receiver_name_var.get()),
        ]

        for label, value in info_items:
            c.drawString(60, y_pos, f"{label}: {value}")
            y_pos -= line_height

        # 備註
        if self.notes_var.get().strip():
            c.drawString(60, y_pos, f"Notes / 備註: {self.notes_var.get()}")
            y_pos -= line_height

        y_pos -= 10

        # 商品標題
        c.setFont(self.font_name, font_size_header)
        c.drawString(60, y_pos, "Items / 商品明細:")
        y_pos -= 30

        # 表格標題
        c.setFont(self.font_name, font_size_content)
        c.drawString(60, y_pos, "Item Code / 商品編號")
        c.drawString(220, y_pos, "Description / 商品描述")
        c.drawString(650, y_pos, "Qty / 數量")

        # 畫線
        c.line(60, y_pos - 5, width - 60, y_pos - 5)
        y_pos -= 25

        # 商品項目
        c.setFont(self.font_name, font_size_small)
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            code, desc, qty = values[0], values[1], values[2]

            # 檢查是否需要新頁面
            if y_pos < 100:
                c.showPage()
                y_pos = height - 80
                c.setFont(self.font_name, font_size_small)

            c.drawString(60, y_pos, str(code))

            # 處理長描述
            max_desc_length = 60
            if len(desc) > max_desc_length:
                desc = desc[:max_desc_length] + "..."
            c.drawString(220, y_pos, desc)

            c.drawString(650, y_pos, str(qty))
            y_pos -= 20

        # 總計
        total_items = len(self.tree.get_children())
        y_pos -= 20
        c.line(60, y_pos, width - 60, y_pos)
        y_pos -= 20
        c.setFont(self.font_name, font_size_content)
        c.drawString(60, y_pos, f"Total Items / 總項目數: {total_items}")

        # 頁腳
        c.setFont(self.font_name, font_size_small)
        footer_text = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        c.drawString(60, 30, footer_text)

        c.save()

    def open_file(self, filepath):
        """開啟檔案"""
        try:
            if platform.system() == "Windows":
                os.startfile(filepath)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", filepath])
            else:  # Linux
                subprocess.run(["xdg-open", filepath])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟檔案: {e}")

    def open_folder(self, path):
        """開啟資料夾"""
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", path])
            else:  # Linux
                subprocess.run(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟資料夾: {e}")


def main():
    """主程式"""
    root = tk.Tk()
    app = PDFGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
