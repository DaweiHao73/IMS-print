import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
import os
import subprocess
import platform
from datetime import datetime

class PDFGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF文件生成器 - Transfer Document Generator")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 設定樣式
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        self.chinese_font = self.setup_chinese_font()
        
    def setup_chinese_font(self):
        """設定中文字體"""
        try:
            if platform.system() == "Windows":
                font_paths = [
                    "C:\\Windows\\Fonts\\msjh.ttc",
                    "C:\\Windows\\Fonts\\simsun.ttc",
                    "C:\\Windows\\Fonts\\kaiu.ttf",
                    "C:\\Windows\\Fonts\\mingliu.ttc"
                ]
                
                for font_path in font_paths:
                    if os.path.exists(font_path):
                        pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                        # 註冊粗體版本
                        pdfmetrics.registerFont(TTFont('ChineseFontBold', font_path))
                        return 'ChineseFont'
            
            return 'Helvetica'
        except:
            return 'Helvetica'
    
    def setup_ui(self):
        # 主標題
        title_frame = ttk.Frame(self.root)
        title_frame.pack(pady=10)
        
        title_label = ttk.Label(title_frame, text="PDF文件生成器", font=('Arial', 16, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="Transfer Document Generator", font=('Arial', 10))
        subtitle_label.pack()
        
        # 主要內容框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 輸入欄位
        self.create_input_fields(main_frame)
        
        # 按鈕區域
        self.create_buttons(main_frame)
        
        # 批量處理區域
        self.create_batch_section(main_frame)
        
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
        
    def create_buttons(self, parent):
        # 按鈕框架
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 生成PDF按鈕
        generate_btn = ttk.Button(button_frame, text="生成PDF", command=self.generate_pdf)
        generate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空表單按鈕
        clear_btn = ttk.Button(button_frame, text="清空表單", command=self.clear_form)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 選擇保存位置按鈕
        self.save_path_var = tk.StringVar(value=os.getcwd())
        path_btn = ttk.Button(button_frame, text="選擇保存位置", command=self.choose_save_path)
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
        path = filedialog.askdirectory(initialdir=self.save_path_var.get())
        if path:
            self.save_path_var.set(path)
            self.path_label.config(text=path)
    
    def clear_form(self):
        self.date_var.set(datetime.now().strftime("%Y/%m/%d"))
        self.sender_store_var.set("")
        self.sender_name_var.set("")
        self.receiver_store_var.set("")
        self.receiver_name_var.set("")
    
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
        # 計算標籤文字寬度並在後面繪製粗體數據
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
        
        # 裝飾邊框
        c.rect(40, 40, width - 80, height - 80, stroke=1, fill=0)
        
        # 簽名區域
        separator_y = height - 340
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
    
    def generate_pdf(self):
        if not self.validate_inputs():
            return
        
        data = {
            'date': self.date_var.get(),
            'sender_store': self.sender_store_var.get(),
            'sender_name': self.sender_name_var.get(),
            'receiver_store': self.receiver_store_var.get(),
            'receiver_name': self.receiver_name_var.get()
        }
        
        filename = f"文件_{data['date'].replace('/', '_')}_{data['sender_store']}_to_{data['receiver_store']}.pdf"
        filepath = os.path.join(self.save_path_var.get(), filename)
        
        try:
            self.create_pdf_document(data, filepath)
            
            result = messagebox.askyesnocancel("成功", f"PDF文件已生成: {filename}\n\n是否要開啟文件？\n(取消=不開啟，是=開啟，否=開啟資料夾)")
            
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
                    'receiver_name': values[4]
                }
                
                filename = f"文件_{i}_{data['date'].replace('/', '_')}_{data['sender_store']}_to_{data['receiver_store']}.pdf"
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
        try:
            if platform.system() == "Windows":
                os.startfile(filepath)
            elif platform.system() == "Darwin":
                subprocess.run(["open", filepath])
            else:
                subprocess.run(["xdg-open", filepath])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟文件: {str(e)}")
    
    def open_folder(self, folder_path):
        try:
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", folder_path])
            else:
                subprocess.run(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟資料夾: {str(e)}")

def main():
    root = tk.Tk()
    app = PDFGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()