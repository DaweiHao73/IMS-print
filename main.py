import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import platform
import subprocess
from datetime import datetime
import json


class PDFGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF文件生成器 - Transfer Document Generator")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 設置最小視窗大小
        self.root.minsize(800, 600)
        
        # 初始化字體
        self.setup_fonts()
        
        # 載入IMS數據
        self.load_ims_data()
        
        # 設置UI
        self.setup_ui()
        
        # 強制刷新顯示
        self.root.update_idletasks()
    
    def setup_fonts(self):
        """設置中文字體 - 支援 Windows 和 macOS"""
        try:
            system = platform.system()
            font_found = False
            
            if system == "Windows":
                # Windows 字體路徑
                windows_fonts = [
                    "C:/Windows/Fonts/msjh.ttc",     # 微軟正黑體
                    "C:/Windows/Fonts/msyh.ttc",     # 微軟雅黑
                    "C:/Windows/Fonts/simhei.ttf",   # 黑體
                    "C:/Windows/Fonts/simsun.ttc",   # 宋體
                    "C:/Windows/Fonts/kaiti.ttf",    # 楷體
                ]
                
                for font_path in windows_fonts:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                            # 對於 Windows，我們使用同一個字體文件但嘗試粗體變體
                            pdfmetrics.registerFont(TTFont('ChineseFontBold', font_path))
                            self.chinese_font = 'ChineseFont'
                            font_found = True
                            print(f"使用字體: {font_path}")
                            break
                        except Exception as e:
                            print(f"字體載入失敗 {font_path}: {e}")
                            continue
            
            elif system == "Darwin":  # macOS
                # macOS 字體路徑
                macos_fonts = [
                    "/System/Library/Fonts/PingFang.ttc",              # 蘋方
                    "/System/Library/Fonts/Helvetica.ttc",             # Helvetica
                    "/System/Library/Fonts/Supplemental/Songti.ttc",   # 宋體
                    "/System/Library/Fonts/Supplemental/Kaiti.ttc",    # 楷體
                    "/Library/Fonts/Microsoft/Microsoft JhengHei.ttf", # 微軟正黑體（如果有安裝）
                ]
                
                for font_path in macos_fonts:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                            pdfmetrics.registerFont(TTFont('ChineseFontBold', font_path))
                            self.chinese_font = 'ChineseFont'
                            font_found = True
                            print(f"使用字體: {font_path}")
                            break
                        except Exception as e:
                            print(f"字體載入失敗 {font_path}: {e}")
                            continue
            
            else:  # Linux 或其他系統
                linux_fonts = [
                    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
                ]
                
                for font_path in linux_fonts:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                            pdfmetrics.registerFont(TTFont('ChineseFontBold', font_path))
                            self.chinese_font = 'ChineseFont'
                            font_found = True
                            print(f"使用字體: {font_path}")
                            break
                        except Exception as e:
                            print(f"字體載入失敗 {font_path}: {e}")
                            continue
            
            if not font_found:
                print("未找到合適的中文字體，使用 Helvetica")
                self.chinese_font = 'Helvetica'
                
        except Exception as e:
            print(f"字體設置發生錯誤: {e}")
            self.chinese_font = 'Helvetica'
    
    def load_ims_data(self):
        """載入IMS數據 - 支援不同路徑格式"""
        self.ims_data = {}
        try:
            # 獲取腳本所在目錄
            script_dir = os.path.dirname(os.path.abspath(__file__))
            json_file = os.path.join(script_dir, 'ims_list.json')
            
            # 如果同目錄下沒有，嘗試當前工作目錄
            if not os.path.exists(json_file):
                json_file = 'ims_list.json'
            
            if os.path.exists(json_file):
                # 使用 UTF-8 編碼確保跨平台兼容
                with open(json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for item in data:
                        if 'Item No' in item and 'Item Description' in item:
                            self.ims_data[item['Item No']] = item['Item Description']
                print(f"成功載入 {len(self.ims_data)} 筆 IMS 數據")
            else:
                print("未找到 ims_list.json 文件，物品查詢功能將無法使用")
                
        except Exception as e:
            print(f"載入IMS數據時發生錯誤: {e}")
            messagebox.showwarning("警告", f"載入IMS數據失敗: {e}\n物品查詢功能將無法使用")
    
    def lookup_description(self):
        """查詢物品描述"""
        article_no = self.article_entry.get().strip()
        if not article_no:
            messagebox.showwarning("警告", "請輸入Article No")
            return
        
        if article_no in self.ims_data:
            self.description_var.set(self.ims_data[article_no])
        else:
            self.description_var.set("未找到相關描述")
    
    def add_item_to_list(self):
        """添加物品到清單"""
        article_no = self.article_entry.get().strip()
        description = self.description_var.get()
        quantity = self.quantity_entry.get().strip()
        
        if not article_no:
            messagebox.showwarning("警告", "請輸入Article No")
            return
        if not quantity:
            messagebox.showwarning("警告", "請輸入數量")
            return
        
        # 添加到物品清單
        values = (article_no, description, quantity)
        self.items_tree.insert('', tk.END, values=values)
        
        # 清空輸入欄位
        self.article_entry.delete(0, tk.END)
        self.description_var.set("")
        self.quantity_entry.delete(0, tk.END)
    
    def remove_item_from_list(self):
        """從清單移除物品"""
        selected = self.items_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要移除的物品")
            return
        
        for item in selected:
            self.items_tree.delete(item)
    
    def clear_items_list(self):
        """清空物品清單"""
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
    
    def setup_ui(self):
        # 清空可能存在的舊內容
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 主標題
        title_frame = ttk.Frame(self.root)
        title_frame.pack(pady=10, fill=tk.X)
        
        title_label = ttk.Label(title_frame, text="PDF文件生成器", font=('Arial', 16, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="Transfer Document Generator", font=('Arial', 10))
        subtitle_label.pack()
        
        # 主要內容框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建筆記本控件來分頁
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 文件資訊頁面
        self.doc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.doc_frame, text="文件資訊")
        
        # 物品清單頁面
        self.items_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.items_frame, text="物品清單")
        
        # 在文件資訊頁面添加內容
        self.create_document_tab(self.doc_frame)
        
        # 在物品清單頁面添加內容
        self.create_items_tab(self.items_frame)
        
        # 按鈕區域（共用）
        self.create_main_buttons(main_frame)
        
        # 強制重新佈局
        self.root.update_idletasks()
        print("UI 設置完成")
    
    def create_document_tab(self, parent):
        """創建文件資訊標籤頁"""
        # 輸入欄位
        self.create_input_fields(parent)
        
        # 批量處理區域
        self.create_batch_section(parent)
    
    def create_items_tab(self, parent):
        """創建物品清單標籤頁"""
        # 物品輸入區域
        items_input_frame = ttk.LabelFrame(parent, text="物品資訊", padding="10")
        items_input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Article No 輸入欄位
        ttk.Label(items_input_frame, text="Article No:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.article_entry = ttk.Entry(items_input_frame, width=30)
        self.article_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 查詢按鈕
        lookup_btn = ttk.Button(items_input_frame, text="查詢描述", command=self.lookup_description)
        lookup_btn.grid(row=0, column=2, padx=(10, 0), pady=2)
        
        # Description 顯示欄位（唯讀）
        ttk.Label(items_input_frame, text="Description:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.description_var = tk.StringVar()
        self.description_entry = ttk.Entry(items_input_frame, textvariable=self.description_var, width=80, state="readonly")
        self.description_entry.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 數量輸入欄位
        ttk.Label(items_input_frame, text="Quantity:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.quantity_entry = ttk.Entry(items_input_frame, width=15)
        self.quantity_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 添加物品按鈕
        add_item_btn = ttk.Button(items_input_frame, text="添加物品", command=self.add_item_to_list)
        add_item_btn.grid(row=2, column=2, padx=(10, 0), pady=2)
        
        # 物品清單區域
        items_list_frame = ttk.LabelFrame(parent, text="物品清單", padding="10")
        items_list_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 物品清單
        items_columns = ('Article No', 'Description', 'Quantity')
        self.items_tree = ttk.Treeview(items_list_frame, columns=items_columns, show='headings', height=8)
        
        self.items_tree.heading('Article No', text='Article No')
        self.items_tree.heading('Description', text='Description')
        self.items_tree.heading('Quantity', text='Quantity')
        
        self.items_tree.column('Article No', width=100)
        self.items_tree.column('Description', width=400)
        self.items_tree.column('Quantity', width=80)
        
        # 滾動條
        items_scrollbar = ttk.Scrollbar(items_list_frame, orient=tk.VERTICAL, command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=items_scrollbar.set)
        
        self.items_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        items_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 物品操作按鈕
        items_btn_frame = ttk.Frame(parent)
        items_btn_frame.pack(fill=tk.X, pady=5)
        
        remove_item_btn = ttk.Button(items_btn_frame, text="移除選中物品", command=self.remove_item_from_list)
        remove_item_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_items_btn = ttk.Button(items_btn_frame, text="清空物品清單", command=self.clear_items_list)
        clear_items_btn.pack(side=tk.LEFT)
    
    def create_input_fields(self, parent):
        # 輸入欄位框架
        input_frame = ttk.LabelFrame(parent, text="文件資訊", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 日期
        ttk.Label(input_frame, text="日期:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        date_entry = ttk.Entry(input_frame, textvariable=self.date_var, width=20)
        date_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 寄出店別
        ttk.Label(input_frame, text="寄出店別:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.sender_store_var = tk.StringVar()
        sender_store_entry = ttk.Entry(input_frame, textvariable=self.sender_store_var, width=30)
        sender_store_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 寄件人
        ttk.Label(input_frame, text="寄件人:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.sender_name_var = tk.StringVar()
        sender_name_entry = ttk.Entry(input_frame, textvariable=self.sender_name_var, width=30)
        sender_name_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 收件店別
        ttk.Label(input_frame, text="收件店別:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.receiver_store_var = tk.StringVar()
        receiver_store_entry = ttk.Entry(input_frame, textvariable=self.receiver_store_var, width=30)
        receiver_store_entry.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 收件人
        ttk.Label(input_frame, text="收件人:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.receiver_name_var = tk.StringVar()
        receiver_name_entry = ttk.Entry(input_frame, textvariable=self.receiver_name_var, width=30)
        receiver_name_entry.grid(row=4, column=1, sticky=tk.W, padx=(10, 0), pady=2)
    
    def create_main_buttons(self, parent):
        # 主按鈕框架
        main_button_frame = ttk.Frame(parent)
        main_button_frame.pack(fill=tk.X, pady=10)
        
        # 生成PDF按鈕
        generate_btn = ttk.Button(main_button_frame, text="生成PDF", command=self.generate_pdf)
        generate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空表單按鈕
        clear_btn = ttk.Button(main_button_frame, text="清空表單", command=self.clear_form)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 選擇保存位置按鈕
        self.save_path_var = tk.StringVar(value=os.path.expanduser("~"))  # 使用用戶主目錄作為預設
        path_btn = ttk.Button(main_button_frame, text="選擇保存位置", command=self.choose_save_path)
        path_btn.pack(side=tk.LEFT)
        
        # 顯示當前保存路徑
        path_frame = ttk.Frame(parent)
        path_frame.pack(fill=tk.X, pady=5)
        ttk.Label(path_frame, text="保存位置:").pack(side=tk.LEFT)
        self.path_label = ttk.Label(path_frame, text=self.save_path_var.get(), foreground="blue")
        self.path_label.pack(side=tk.LEFT, padx=(5, 0))
    
    def create_batch_section(self, parent):
        # 批量處理框架
        batch_frame = ttk.LabelFrame(parent, text="批量處理", padding="10")
        batch_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 批量列表
        columns = ('日期', '寄出店別', '寄件人', '收件店別', '收件人')
        self.batch_tree = ttk.Treeview(batch_frame, columns=columns, show='headings', height=6)
        
        for col in columns:
            self.batch_tree.heading(col, text=col)
            self.batch_tree.column(col, width=100)
        
        # 滾動條
        scrollbar = ttk.Scrollbar(batch_frame, orient=tk.VERTICAL, command=self.batch_tree.yview)
        self.batch_tree.configure(yscrollcommand=scrollbar.set)
        
        self.batch_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 批量操作按鈕
        batch_btn_frame = ttk.Frame(parent)
        batch_btn_frame.pack(fill=tk.X, pady=5)
        
        add_to_batch_btn = ttk.Button(batch_btn_frame, text="添加到批量列表", command=self.add_to_batch)
        add_to_batch_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        remove_from_batch_btn = ttk.Button(batch_btn_frame, text="從列表移除", command=self.remove_from_batch)
        remove_from_batch_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        generate_batch_btn = ttk.Button(batch_btn_frame, text="批量生成PDF", command=self.generate_batch_pdf)
        generate_batch_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_batch_btn = ttk.Button(batch_btn_frame, text="清空列表", command=self.clear_batch)
        clear_batch_btn.pack(side=tk.LEFT)
    
    def choose_save_path(self):
        """選擇保存路徑 - 跨平台兼容"""
        try:
            # 使用當前保存路徑作為初始目錄
            initial_dir = self.save_path_var.get()
            
            # 確保初始目錄存在
            if not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")  # 使用用戶主目錄
            
            path = filedialog.askdirectory(
                title="選擇保存位置",
                initialdir=initial_dir
            )
            
            if path:
                # 標準化路徑格式（跨平台）
                normalized_path = os.path.normpath(path)
                self.save_path_var.set(normalized_path)
                
                # 限制顯示路徑長度
                display_path = normalized_path
                if len(display_path) > 50:
                    display_path = "..." + display_path[-47:]
                
                self.path_label.config(text=display_path)
                
        except Exception as e:
            messagebox.showerror("錯誤", f"選擇路徑時發生錯誤: {str(e)}")
    
    def clear_form(self):
        self.date_var.set(datetime.now().strftime("%Y/%m/%d"))
        self.sender_store_var.set("")
        self.sender_name_var.set("")
        self.receiver_store_var.set("")
        self.receiver_name_var.set("")
        # 清空物品清單
        self.clear_items_list()
    
    def validate_inputs(self):
        if not self.sender_store_var.get().strip():
            messagebox.showerror("錯誤", "請輸入寄出店別")
            return False
        if not self.sender_name_var.get().strip():
            messagebox.showerror("錯誤", "請輸入寄件人")
            return False
        if not self.receiver_store_var.get().strip():
            messagebox.showerror("錯誤", "請輸入收件店別")
            return False
        if not self.receiver_name_var.get().strip():
            messagebox.showerror("錯誤", "請輸入收件人")
            return False
        return True
    
    def create_pdf_document(self, data, filename):
        """創建PDF文件"""
        c = canvas.Canvas(filename, pagesize=landscape(A4))
        width, height = landscape(A4)
        
        # 標題
        c.setFont("Helvetica-Bold", 20)
        title_x = (width - 300) / 2
        c.drawString(title_x, height - 60, "Transfer Document / 調貨單")
        
        # 內容 - 標籤部分使用普通字體，數據部分使用粗體
        y_position = height - 120
        line_height = 40
        left_margin = 80
        
        # 獲取適合的粗體字體
        if self.chinese_font == 'ChineseFont':
            bold_font = 'ChineseFontBold'
        else:
            bold_font = 'Helvetica-Bold'
        
        # 日期
        c.setFont(self.chinese_font, 14)
        c.drawString(left_margin, y_position, "日期 Date: ")
        label_width = c.stringWidth("日期 Date: ", self.chinese_font, 14)
        c.setFont(bold_font, 14)
        c.drawString(left_margin + label_width, y_position, data['date'])
        y_position -= line_height
        
        # 寄出店別
        c.setFont(self.chinese_font, 14)
        c.drawString(left_margin, y_position, "寄出店別 From Store: ")
        label_width = c.stringWidth("寄出店別 From Store: ", self.chinese_font, 14)
        c.setFont(bold_font, 14)
        c.drawString(left_margin + label_width, y_position, data['sender_store'])
        y_position -= line_height
        
        # 寄件人
        c.setFont(self.chinese_font, 14)
        c.drawString(left_margin, y_position, "寄件人 Sender: ")
        label_width = c.stringWidth("寄件人 Sender: ", self.chinese_font, 14)
        c.setFont(bold_font, 14)
        c.drawString(left_margin + label_width, y_position, data['sender_name'])
        y_position -= line_height
        
        # 收件店別
        c.setFont(self.chinese_font, 14)
        c.drawString(left_margin, y_position, "收件店別 To Store: ")
        label_width = c.stringWidth("收件店別 To Store: ", self.chinese_font, 14)
        c.setFont(bold_font, 14)
        c.drawString(left_margin + label_width, y_position, data['receiver_store'])
        y_position -= line_height
        
        # 收件人
        c.setFont(self.chinese_font, 14)
        c.drawString(left_margin, y_position, "收件人 Receiver: ")
        label_width = c.stringWidth("收件人 Receiver: ", self.chinese_font, 14)
        c.setFont(bold_font, 14)
        c.drawString(left_margin + label_width, y_position, data['receiver_name'])
        y_position -= line_height * 1.5
        
        # 物品清單
        if data.get('items'):
            c.setFont(self.chinese_font, 16)
            c.drawString(left_margin, y_position, "物品清單 Items List:")
            y_position -= 30
            
            # 表格標題
            c.setFont(self.chinese_font, 12)
            c.drawString(left_margin, y_position, "Article No")
            c.drawString(left_margin + 120, y_position, "Description")
            c.drawString(left_margin + 500, y_position, "Quantity")
            y_position -= 5
            
            # 畫線分隔
            c.line(left_margin, y_position, width - 80, y_position)
            y_position -= 20
            
            # 物品詳細
            for item in data['items']:
                if y_position < 150:  # 如果空間不夠，換頁
                    c.showPage()
                    y_position = height - 80
                
                c.setFont("Helvetica", 10)
                c.drawString(left_margin, y_position, item['article_no'])
                
                # 處理長描述，可能需要換行
                description = item['description']
                if len(description) > 40:
                    description = description[:40] + "..."
                c.drawString(left_margin + 120, y_position, description)
                
                c.drawString(left_margin + 500, y_position, item['quantity'])
                y_position -= 20
        
        # 裝飾邊框
        c.rect(40, 40, width - 80, height - 80, stroke=1, fill=0)
        
        # 簽名區域
        separator_y = 180
        c.line(60, separator_y, width - 60, separator_y)
        
        signature_y = separator_y - 60
        left_col_x = 80
        right_col_x = width / 2 + 50
        
        # 簽名區域使用普通字體
        c.setFont(self.chinese_font, 14)
        
        # 左欄：寄件人簽名
        c.drawString(left_col_x, signature_y, "寄件人簽名 Sender Signature:")
        c.line(left_col_x + 220, signature_y - 5, right_col_x - 30, signature_y - 5)
        c.drawString(left_col_x, signature_y - 40, "日期 Date:")
        c.line(left_col_x + 80, signature_y - 45, left_col_x + 200, signature_y - 45)
        
        # 右欄：收件人簽名
        c.drawString(right_col_x, signature_y, "收件人簽名 Receiver Signature:")
        c.line(right_col_x + 220, signature_y - 5, width - 80, signature_y - 5)
        c.drawString(right_col_x, signature_y - 40, "日期 Date:")
        c.line(right_col_x + 80, signature_y - 45, right_col_x + 200, signature_y - 45)
        
        c.save()
    
    def get_items_data(self):
        """獲取物品清單數據"""
        items = []
        for item in self.items_tree.get_children():
            values = self.items_tree.item(item)['values']
            items.append({
                'article_no': values[0],
                'description': values[1],
                'quantity': values[2]
            })
        return items
    
    def generate_pdf(self):
        if not self.validate_inputs():
            return
        
        items_data = self.get_items_data()
        
        data = {
            'date': self.date_var.get(),
            'sender_store': self.sender_store_var.get(),
            'sender_name': self.sender_name_var.get(),
            'receiver_store': self.receiver_store_var.get(),
            'receiver_name': self.receiver_name_var.get(),
            'items': items_data
        }
        
        filename = f"調貨單_{data['date'].replace('/', '_')}_{data['sender_store']}_to_{data['receiver_store']}.pdf"
        filepath = os.path.join(self.save_path_var.get(), filename)
        
        try:
            self.create_pdf_document(data, filepath)
            
            result = messagebox.askyesnocancel("成功",
                                               f"PDF文件已生成: {filename}\n\n是否要開啟文件？\n(取消=不開啟，是=開啟，否=開啟資料夾)")
            
            if result is True:  # 開啟文件
                self.open_pdf(filepath)
            elif result is False:  # 開啟資料夾
                self.open_folder(self.save_path_var.get())
        
        except Exception as e:
            messagebox.showerror("錯誤", f"生成PDF時發生錯誤: {str(e)}")
    
    def add_to_batch(self):
        if not self.validate_inputs():
            return
        
        values = (
            self.date_var.get(),
            self.sender_store_var.get(),
            self.sender_name_var.get(),
            self.receiver_store_var.get(),
            self.receiver_name_var.get()
        )
        
        self.batch_tree.insert('', tk.END, values=values)
        self.clear_form()
        messagebox.showinfo("成功", "已添加到批量列表")
    
    def remove_from_batch(self):
        selected = self.batch_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇要移除的項目")
            return
        
        for item in selected:
            self.batch_tree.delete(item)
    
    def clear_batch(self):
        for item in self.batch_tree.get_children():
            self.batch_tree.delete(item)
    
    def generate_batch_pdf(self):
        items = self.batch_tree.get_children()
        if not items:
            messagebox.showwarning("警告", "批量列表為空")
            return
        
        generated_files = []
        
        try:
            for i, item in enumerate(items, 1):
                values = self.batch_tree.item(item)['values']
                data = {
                    'date': values[0],
                    'sender_store': values[1],
                    'sender_name': values[2],
                    'receiver_store': values[3],
                    'receiver_name': values[4],
                    'items': []  # 批量生成時暫不包含物品清單
                }
                
                filename = f"調貨單_{i}_{data['date'].replace('/', '_')}_{data['sender_store']}_to_{data['receiver_store']}.pdf"
                filepath = os.path.join(self.save_path_var.get(), filename)
                
                self.create_pdf_document(data, filepath)
                generated_files.append(filename)
            
            messagebox.showinfo("成功", f"批量生成完成！\n共生成了 {len(generated_files)} 個文件")
            
            # 詢問是否開啟資料夾
            if messagebox.askyesno("完成", "是否要開啟保存資料夾？"):
                self.open_folder(self.save_path_var.get())
        
        except Exception as e:
            messagebox.showerror("錯誤", f"批量生成時發生錯誤: {str(e)}")
    
    def open_pdf(self, filepath):
        """跨平台開啟PDF文件"""
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(filepath)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", filepath], check=True)
            else:  # Linux 和其他 Unix-like 系統
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟文件: {str(e)}")
    
    def open_folder(self, folder_path):
        """跨平台開啟資料夾"""
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(folder_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", folder_path], check=True)
            else:  # Linux 和其他 Unix-like 系統
                subprocess.run(["xdg-open", folder_path], check=True)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟資料夾: {str(e)}")


def main():
    # 設置環境變數來消除警告
    import os
    
    # 消除 macOS 的 Tk 廢棄警告
    if platform.system() == "Darwin":
        os.environ['TK_SILENCE_DEPRECATION'] = '1'
    
    try:
        root = tk.Tk()
        
        # 跨平台的視窗設置
        if platform.system() == "Darwin":  # macOS
            # macOS 特定設置
            root.lift()
            root.attributes('-topmost', True)
            root.after_idle(root.attributes, '-topmost', False)
        elif platform.system() == "Windows":  # Windows
            # Windows 特定設置
            root.state('normal')
            root.lift()
            root.focus_force()
        
        # 確保視窗在螢幕中央（跨平台）
        root.update_idletasks()
        
        # 獲取螢幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # 計算視窗位置
        window_width = 900
        window_height = 700
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # 確保視窗不會超出螢幕邊界
        x = max(0, min(x, screen_width - window_width))
        y = max(0, min(y, screen_height - window_height))
        
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 創建應用程式實例
        app = PDFGeneratorApp(root)
        
        # 強制重新繪製
        root.update()
        root.deiconify()  # 確保視窗顯示
        
        print("=" * 50)
        print("PDF文件生成器已啟動")
        print(f"運行平台: {platform.system()}")
        print("如果看不到GUI視窗，請檢查:")
        if platform.system() == "Darwin":
            print("- 按 Cmd+Tab 查看是否在其他桌面")
            print("- 檢查 Mission Control (F3)")
        elif platform.system() == "Windows":
            print("- 檢查工作列是否有程式圖標")
            print("- 按 Alt+Tab 查看已打開的視窗")
        print("=" * 50)
        
        # 啟動主循環
        root.mainloop()
        
    except Exception as e:
        print(f"程式啟動失敗: {e}")
        import traceback
        traceback.print_exc()
        
        # 嘗試顯示錯誤對話框
        try:
            root = tk.Tk()
            root.withdraw()  # 隱藏主視窗
            messagebox.showerror("錯誤", f"程式啟動失敗:\n{e}")
        except:
            print("無法顯示錯誤對話框")


if __name__ == "__main__":
    main()