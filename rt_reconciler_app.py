import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime
import os
import threading

class ReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RT Reconciler")
        self.root.geometry("500x350")
        self.root.resizable(False, False)
        self.simple_file = tk.StringVar()
        self.rt_file = tk.StringVar()
        self.create_widgets()
    
    def create_widgets(self):
        tk.Label(self.root, text="RT vs Simple Reconciler", font=("Arial", 16, "bold"), pady=20).pack()
        
        frame1 = tk.Frame(self.root, pady=10)
        frame1.pack(fill="x", padx=20)
        tk.Label(frame1, text="Simple Workbook:").pack(anchor="w")
        f1 = tk.Frame(frame1)
        f1.pack(fill="x", pady=5)
        tk.Entry(f1, textvariable=self.simple_file, width=50).pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="Browse...", command=self.browse_simple).pack(side="right", padx=5)
        
        frame2 = tk.Frame(self.root, pady=10)
        frame2.pack(fill="x", padx=20)
        tk.Label(frame2, text="RT Export:").pack(anchor="w")
        f2 = tk.Frame(frame2)
        f2.pack(fill="x", pady=5)
        tk.Entry(f2, textvariable=self.rt_file, width=50).pack(side="left", fill="x", expand=True)
        tk.Button(f2, text="Browse...", command=self.browse_rt).pack(side="right", padx=5)
        
        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=400)
        self.progress.pack(pady=20)
        
        self.status_var = tk.StringVar(value="Select files and click Reconcile")
        tk.Label(self.root, textvariable=self.status_var).pack(pady=10)
        
        tk.Button(self.root, text="Reconcile", font=("Arial", 12, "bold"), width=20, height=2, command=self.start_reconcile).pack(pady=10)
    
    def browse_simple(self):
        f = filedialog.askopenfilename(title="Select Simple Workbook", filetypes=[("Excel", "*.xlsx *.xls")])
        if f: self.simple_file.set(f)
    
    def browse_rt(self):
        f = filedialog.askopenfilename(title="Select RT Export", filetypes=[("Excel", "*.xlsx *.xls")])
        if f: self.rt_file.set(f)
    
    def start_reconcile(self):
        if not self.simple_file.get() or not self.rt_file.get():
            messagebox.showerror("Error", "Please select both files")
            return
        self.progress.start()
        self.status_var.set("Processing...")
        threading.Thread(target=self.run_reconcile).start()
    
    def run_reconcile(self):
        try:
            output = self.reconcile(self.simple_file.get(), self.rt_file.get())
            self.root.after(0, lambda: self.on_complete(output))
        except Exception as e:
            self.root.after(0, lambda: self.on_error(str(e)))
    
    def on_complete(self, output_file):
        self.progress.stop()
        self.status_var.set("Complete!")
        if messagebox.askyesno("Success", f"Saved to:\n{output_file}\n\nOpen now?"):
            os.startfile(output_file)
    
    def on_error(self, msg):
        self.progress.stop()
        self.status_var.set("Error")
        messagebox.showerror("Error", msg)
    
    def reconcile(self, simple_file, rt_file):
        simple_pivot = pd.read_excel(simple_file, sheet_name='Sheet1', skiprows=1)
        simple_pivot.columns = ['IET #', 'SIMPLE']
        simple_pivot = simple_pivot.dropna(subset=['IET #'])
        simple_pivot = simple_pivot[~simple_pivot['IET #'].astype(str).str.lower().str.contains('grand total|blank|row labels', na=False)]
        simple_pivot['SIMPLE'] = pd.to_numeric(simple_pivot['SIMPLE'], errors='coerce').fillna(0).astype(int)
        detail_df = pd.read_excel(simple_file, sheet_name='IE Tire')
        
        rt_df = pd.read_excel(rt_file, sheet_name=0)
        if 'RT' in rt_df.columns:
            rt_data = rt_df[[rt_df.columns[0], 'RT']].copy()
        else:
            rt_data = rt_df.iloc[:, [0, -1]].copy()
        rt_data.columns = ['IET #', 'RT']
        rt_data = rt_data.dropna(subset=['IET #'])
        rt_data['RT'] = pd.to_numeric(rt_data['RT'], errors='coerce').fillna(0).astype(int)
        
        merged = simple_pivot.merge(rt_data, on='IET #', how='left')
        merged['RT'] = merged['RT'].fillna(0).astype(int)
        merged['DIFF'] = merged['SIMPLE'] - merged['RT']
        
        reconciled = merged[merged['DIFF'] == 0].copy()
        discrepancies = merged[merged['DIFF'] != 0].sort_values('DIFF', key=abs, ascending=False)
        not_in_rt = discrepancies[discrepancies['RT'] == 0].copy()
        
        # Build dict of IET # -> DIFF amount (how many UNITS to highlight per SKU)
        diff_dict = {}
        for _, row in discrepancies.iterrows():
            iet = str(row['IET #'])
            diff = int(row['DIFF'])
            if diff > 0:  # Only highlight if Simple has more than RT
                diff_dict[iet] = diff
        
        output_file = os.path.join(os.path.dirname(simple_file), f"Reconciled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame({'Metric': ['Total SKUs', 'Reconciled', 'Discrepancies', 'Not in RT', '', 'SIMPLE Total', 'RT Total', 'Variance'], 'Value': [len(merged), len(reconciled), len(discrepancies), len(not_in_rt), '', merged['SIMPLE'].sum(), merged['RT'].sum(), merged['DIFF'].sum()]}).to_excel(writer, sheet_name='Summary', index=False)
            discrepancies.to_excel(writer, sheet_name='Discrepancies', index=False)
            reconciled.to_excel(writer, sheet_name='Reconciled', index=False)
            not_in_rt.to_excel(writer, sheet_name='Not in RT', index=False)
            detail_df.to_excel(writer, sheet_name='IE Tire Detail', index=False)
            merged.to_excel(writer, sheet_name='Full Comparison', index=False)
        
        # Apply formatting and highlight based on return_qty
        self.format_and_highlight(output_file, diff_dict)
        
        return output_file
    
    def format_and_highlight(self, file_path, diff_dict):
        wb = load_workbook(file_path)
        
        header_fill = PatternFill('solid', fgColor='4472C4')
        header_font = Font(color='FFFFFF', bold=True)
        green = PatternFill('solid', fgColor='C6EFCE')
        yellow = PatternFill('solid', fgColor='FFEB9C')
        red = PatternFill('solid', fgColor='FFC7CE')
        
        for ws in wb.worksheets:
            # Format headers
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            # Auto-fit columns
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)
            
            # Color code DIFF column in comparison sheets
            if ws.title in ['Discrepancies', 'Reconciled', 'Not in RT', 'Full Comparison']:
                diff_col = None
                for i, cell in enumerate(ws[1], 1):
                    if cell.value == 'DIFF':
                        diff_col = i
                if diff_col:
                    for row in ws.iter_rows(min_row=2):
                        val = row[diff_col-1].value
                        if val == 0:
                            fill = green
                        elif val and val > 0:
                            fill = yellow
                        elif val and val < 0:
                            fill = red
                        else:
                            continue
                        for cell in row[:diff_col]:
                            cell.fill = fill
            
            # Highlight based on return_qty in IE Tire Detail
            if ws.title == 'IE Tire Detail':
                # Find IET # and return_qty columns
                iet_col = None
                qty_col = None
                for i, cell in enumerate(ws[1], 1):
                    if cell.value == 'IET #':
                        iet_col = i
                    if cell.value == 'return_qty':
                        qty_col = i
                
                if iet_col and qty_col:
                    # Track how many UNITS we've highlighted per SKU
                    highlight_qty = {k: 0 for k in diff_dict.keys()}
                    
                    for row in ws.iter_rows(min_row=2):
                        iet_value = str(row[iet_col-1].value or '')
                        row_qty = row[qty_col-1].value or 1
                        try:
                            row_qty = int(row_qty)
                        except:
                            row_qty = 1
                        
                        if iet_value in diff_dict:
                            # Only highlight if we haven't reached the DIFF amount yet
                            remaining = diff_dict[iet_value] - highlight_qty[iet_value]
                            if remaining > 0:
                                for cell in row:
                                    cell.fill = red
                                highlight_qty[iet_value] += row_qty
        
        wb.save(file_path)

if __name__ == '__main__':
    root = tk.Tk()
    ReconcilerApp(root)
    root.mainloop()
