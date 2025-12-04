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
        
        # Build dict of IET # -> DIFF amount
        diff_dict = {}
        for _, row in discrepancies.iterrows():
            iet = str(row['IET #'])
            diff = int(row['DIFF'])
            if diff > 0:
                diff_dict[iet] = diff
        
        # Calculate which rows to highlight (red=full, yellow=partial)
        red_rows, yellow_rows = self.calculate_rows_to_highlight(detail_df, diff_dict)
        
        output_file = os.path.join(os.path.dirname(simple_file), f"Reconciled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame({'Metric': ['Total SKUs', 'Reconciled', 'Discrepancies', 'Not in RT', '', 'SIMPLE Total', 'RT Total', 'Variance', '', 'Red = Full row variance', 'Yellow = Partial (row qty > needed)'], 'Value': [len(merged), len(reconciled), len(discrepancies), len(not_in_rt), '', merged['SIMPLE'].sum(), merged['RT'].sum(), merged['DIFF'].sum(), '', '', '']}).to_excel(writer, sheet_name='Summary', index=False)
            discrepancies.to_excel(writer, sheet_name='Discrepancies', index=False)
            reconciled.to_excel(writer, sheet_name='Reconciled', index=False)
            not_in_rt.to_excel(writer, sheet_name='Not in RT', index=False)
            detail_df.to_excel(writer, sheet_name='IE Tire Detail', index=False)
            merged.to_excel(writer, sheet_name='Full Comparison', index=False)
        
        self.format_and_highlight(output_file, red_rows, yellow_rows)
        
        return output_file
    
    def calculate_rows_to_highlight(self, detail_df, diff_dict):
        """Figure out which rows to highlight. Red=full, Yellow=partial."""
        red_rows = set()
        yellow_rows = set()
        
        for iet, diff_needed in diff_dict.items():
            # Get all rows for this IET #, sorted by return_qty ascending
            iet_rows = detail_df[detail_df['IET #'].astype(str) == iet].copy()
            iet_rows = iet_rows.sort_values('return_qty', ascending=True)
            
            qty_remaining = diff_needed
            
            for idx, row in iet_rows.iterrows():
                if qty_remaining <= 0:
                    break
                    
                row_qty = int(row['return_qty']) if pd.notna(row['return_qty']) else 1
                
                if row_qty <= qty_remaining:
                    # Full row counts - RED
                    red_rows.add(idx)
                    qty_remaining -= row_qty
                else:
                    # Partial row - row has more than we need - YELLOW
                    yellow_rows.add(idx)
                    qty_remaining = 0  # We're done with this SKU
        
        return red_rows, yellow_rows
    
    def format_and_highlight(self, file_path, red_rows, yellow_rows):
        wb = load_workbook(file_path)
        
        header_fill = PatternFill('solid', fgColor='4472C4')
        header_font = Font(color='FFFFFF', bold=True)
        green = PatternFill('solid', fgColor='C6EFCE')
        yellow = PatternFill('solid', fgColor='FFEB9C')
        red = PatternFill('solid', fgColor='FFC7CE')
        
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)
            
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
            
            if ws.title == 'IE Tire Detail':
                for excel_row_num, row in enumerate(ws.iter_rows(min_row=2), start=0):
                    if excel_row_num in red_rows:
                        for cell in row:
                            cell.fill = red
                    elif excel_row_num in yellow_rows:
                        for cell in row:
                            cell.fill = yellow
        
        wb.save(file_path)

if __name__ == '__main__':
    root = tk.Tk()
    ReconcilerApp(root)
    root.mainloop()
