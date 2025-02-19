import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def adjust_to_shift(start_time):
    shifts = [(datetime.strptime("07:30", "%H:%M").time(), datetime.strptime("16:30", "%H:%M").time()),
              (datetime.strptime("19:30", "%H:%M").time(), datetime.strptime("04:30", "%H:%M").time())]
    
    while True:
        start_time_time = start_time.time()
        for shift_start, shift_end in shifts:
            if shift_start <= shift_end:
                if shift_start <= start_time_time <= shift_end:
                    return start_time
            else:
                if start_time_time >= shift_start or start_time_time <= shift_end:
                    return start_time
        start_time -= timedelta(hours=1)

def process_bom():
    input_file = file_entry.get()
    main_part_number = part_entry.get()
    amount = int(amount_entry.get())
    due_datetime = datetime.strptime(due_entry.get(), "%d/%m/%Y %H:%M")
    output_file = "Processed_BOM.xlsx"
    
    if not input_file:
        messagebox.showerror("Error", "Please select a BOM file.")
        return
    
    bom_df = pd.read_excel(input_file)
    required_parts = {}
    
    def find_sub_parts(part_number, qty, due_time):
        sub_parts = bom_df[bom_df['main_partnumber'] == part_number]
        
        for _, row in sub_parts.iterrows():
            sub_part = row['sub_partnumber']
            sub_qty = row['Sub Qty'] * qty
            lead_time = row['Lead Time (sec)']
            start_time = due_time - timedelta(seconds=lead_time)
            start_time = adjust_to_shift(start_time)
            
            if sub_part in required_parts:
                required_parts[sub_part]['quantity'] += sub_qty
                required_parts[sub_part]['start_time'] = min(required_parts[sub_part]['start_time'], start_time)
            else:
                required_parts[sub_part] = {'quantity': sub_qty, 'start_time': start_time}
            
            find_sub_parts(sub_part, sub_qty, start_time)
    
    find_sub_parts(main_part_number, amount, due_datetime)
    result_df = pd.DataFrame([(k, v['quantity'], v['start_time'].strftime('%d/%m/%Y %H:%M')) for k, v in required_parts.items()],
                              columns=['Part Number', 'Total Quantity', 'Start Time'])
    result_df.to_excel(output_file, index=False)
    messagebox.showinfo("Success", f"BOM processed and saved to {output_file}")

root = tk.Tk()
root.title("BOM Processor")
root.geometry("400x300")

tk.Label(root, text="BOM File:").pack()
file_entry = tk.Entry(root, width=50)
file_entry.pack()
tk.Button(root, text="Browse", command=browse_file).pack()

tk.Label(root, text="Main Part Number:").pack()
part_entry = tk.Entry(root)
part_entry.pack()

tk.Label(root, text="Quantity:").pack()
amount_entry = tk.Entry(root)
amount_entry.pack()

tk.Label(root, text="Due Date (dd/mm/yyyy HH:MM):").pack()
due_entry = tk.Entry(root)
due_entry.pack()

tk.Button(root, text="Process BOM", command=process_bom).pack()
root.mainloop()
