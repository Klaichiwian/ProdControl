import requests
import openpyxl
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime, timedelta
import os

API_URL = "http://167.172.68.211/api/load-order"
TOKEN = "afc1d650-024d-4615-bfbf-c01ad42ddbc8"

class ExcelAPIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Fetcher")
        
        self.folder_path = tk.StringVar()
        
        tk.Label(root, text="Select Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(root, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)
        
        tk.Button(root, text="Run", command=self.process_files).grid(row=1, column=0, columnspan=3, pady=10)
        
        self.output_text = tk.Text(root, height=10, width=70)
        self.output_text.grid(row=2, column=0, columnspan=3, padx=5, pady=5)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
    
    def get_month_filename(self):
        return os.path.join(self.folder_path.get(), f"{datetime.today().strftime('%B-%Y')}.xlsx")
    
    def get_order_qty(self, pno, date):
        params = {"token": TOKEN, "dateStart": date, "dateEnd": date, "pno": pno}
        response = requests.get(API_URL, params=params)
        if response.status_code == 200:
            data = response.json()
            return sum(item['o_order_qty'] for item in data[0]['payload'])
        return 0
    
    def modify_dates(self, ws):
        today = datetime.today()
        last_day = (today.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        column_index = 48
        while column_index >= 6:
            ws.cell(row=7, column=column_index, value=last_day.strftime('%Y-%m-%d'))
            last_day -= timedelta(days=1)
            column_index -= 1
    
    def process_files(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Error", "Please select a folder.")
            return
        
        master_file = os.path.join(folder, "master.xlsx")
        if not os.path.exists(master_file):
            messagebox.showerror("Error", "Master file not found in the selected folder.")
            return
        
        new_month_file = self.get_month_filename()
        
        if os.path.exists(new_month_file):
            wb = openpyxl.load_workbook(new_month_file)
        else:
            wb = openpyxl.load_workbook(master_file)
            self.modify_dates(wb.active)
        
        ws = wb.active
        
        for row in range(2, ws.max_row + 1):
            if ws[f"A{row}"].value is not None:
                pno = ws[f"B{row}"].value
                for col in range(6, 49):
                    date_cell = ws.cell(row=7, column=col)
                    if date_cell.value:
                        date_value = date_cell.value
                        if isinstance(date_value, datetime):
                            date_value = date_value.strftime('%Y-%m-%d')
                        
                        total_order_qty = self.get_order_qty(pno, date_value)
                        ws.cell(row=row, column=col).value = total_order_qty
                        self.output_text.insert(tk.END, f"Written {total_order_qty} for {pno} on {date_value}\n")
                        self.output_text.see(tk.END)
        
        if os.path.exists(new_month_file):
            prev_wb = openpyxl.load_workbook(new_month_file)
            prev_ws = prev_wb.active
            for row in range(2, ws.max_row + 1):
                if prev_ws[f"AY{row}"].value == 1:
                    ws[f"E{row}"].value = prev_ws[f"AV{row}"].value
        
        wb.save(new_month_file)
        messagebox.showinfo("Success", f"File updated: {new_month_file}")

root = tk.Tk()
app = ExcelAPIApp(root)
root.mainloop()
