"""
RT vs Simple Tire Reconciliation Tool
Desktop GUI Application
"""

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
        self.root.title("RT Reconciler - Import Export Tire")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # File paths
        self.simple_file = tk.StringVar()
        self.rt_file = tk.StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        # Header
        header = tk.Label(
            self.root, 
            text="RT vs Simple Tire Reconciliation",
            font=("Arial", 16, "bold"),
            pady=20
        )
        header.pack()
        
        # Simple Workbook file selection
        frame1 = tk.Frame(self.root, pady=10)
        frame1.pack(fill="x", padx=20)
        
        tk.Label(frame1, text="Simple Workbook:", font=("Arial", 10)).pack(anchor="w")
        
        file_frame1 = tk.Frame(frame1)
        file_frame1.pack(fill="x", pady=5)
        
        self.simple_entry = tk.Entry(file_frame1, textvariable=self.simple_file, width=50)
        self.simple_entry.pack(side="left", fill="x", expand=True)
        
        tk.Button(file_frame1, text="Browse...", command=self.browse_simple).pack(side="right", padx=5)
        
        # RT Comparison file selection
        frame2 = tk.Frame(self.root, pady=10)
        frame2.pack(fill="x", padx=20)
        
        tk.Label(frame2, text="RT Comparison File:", font=("Arial", 10)).pack(anchor="w")
        
        file_frame2 = tk.Frame(frame2)
        file_frame2.pack(fill="x", pady=5)
        
        self.rt_entry = tk.Entry(file_frame2, textvariable=self.rt_file, width=50)
        self.rt_entry.pack(side="left", fill="x", expand=True)
        
        tk.Button(file_frame2, text="Browse...", command=self.browse_rt).pack(side="right", padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=400)
        self.progress.pack(pady=20)
        
        # Status label
        self.status_var = tk.StringVar(value="Select files and click Reconcile")
        self.status_label = tk.Label(self.root, textvariable=self.status_var, font=("Arial", 10))
        self.status_label.pack(pady=10)
        
        # Reconcile button
        self.reconcile_btn = tk.Button(
            self.root, 
            text="Reconcile", 
            font=("Arial", 12, "bold"),
            bg="#4472C4",
            fg="white",
            width=20,
            height=2,
            command=self.start_reconcile
        )
        self.reconcile_btn.pack(pady=20)
        
        # Footer
        footer = tk.Label(
            self.root,
            text="Import Export Tire Company",
            font=("Arial", 8),
            fg="gray"
        )
        footer.pack(side="bottom", pady=10)
    
    def browse_simple(self):
        filename = filedialog.askopenfilename(
            title="Select Simple Workbook",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.simple_file.set(filename)
    
    def browse_rt(self):
        filename = filedialog.askopenfilename(
            title="Select RT Comparison File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.rt_file.set(filename)
    
    def start_reconcile(self):
        if not self.simple_file.get() or not self.rt_file.get():
            messagebox.showerror("Error", "Please select both files")
            return
        
        # Disable button and start progress
        self.reconcile_btn.config(state="disabled")
        self.progress.start()
        self.status_var.set("Processing...")
        
        # Run in thread to keep UI responsive
        thread = threading.Thread(target=self.run_reconcile)
        thread.start()
    
    def run_reconcile(self):
        try:
            output_file = self.reconcile_simple_workbook(
                self.simple_file.get(),
                self.rt_file.get()
            )
            
            self.root.after(0, lambda: self.on_complete(output_file))
            
        except Exception as e:
            self.root.after(0, lambda: self.on_error(str(e)))
    
    def on_complete(self, output_file):
        self.progress.stop()
        self.reconcile_btn.config(state="normal")
        self.status_var.set(f"Complete! Saved to:\n{os.path.basename(output_file)}")
        
        result = messagebox.askyesno(
            "Success", 
            f"Reconciliation complete!\n\nSaved to:\n{output_file}\n\nOpen the file now?"
        )
        if result:
            os.startfile(output_file)
    
    def on_error(self, error_msg):
        self.progress.stop()
        self.reconcile_btn.config(state="normal")
        self.status_var.set("Error occurred")
        messagebox.showerror("Error", f"Reconciliation failed:\n\n{error_msg}")
    
    def reconcile_simple_workbook(self, simple_file, rt_file):
        """Process Simple_Workbook, compare against RT, and organize into tabs."""
        
        # Load Simple workbook
        simple_pivot = pd.read_excel(simple_file, sheet_name='Sheet1', skiprows=1)
        simple_pivot.columns = ['IET #', 'SIMPLE']
        simple_pivot = simple_pivot.dropna(subset=['IET #'])
        
        # Remove header row, grand total, and blank summary rows
        exclude_patterns = ['Grand Total', 'Row Labels', 'blank']
        mask = simple_pivot['IET #'].astype(str).str.lower().str.contains('|'.join([p.lower() for p in exclude_patterns]), na=False)
        simple_pivot = simple_pivot[~mask]
        
        # Convert SIMPLE to numeric
        simple_pivot['SIMPLE'] = pd.to_numeric(simple_pivot['SIMPLE'], errors='coerce').fillna(0).astype(int)
        
        # Load the detail data
        detail_df = pd.read_excel(simple_file, sheet_name='IE Tire')
        
        # Load RT comparison data
        rt_df = pd.read_excel(rt_file, sheet_name='Sheet1')
        rt_df.columns = ['IET #', 'SIMPLE_RT', 'RT', 'DIFF_RT']
        rt_df = rt_df.dropna(subset=['IET #'])
        
        # Merge Simple with RT data
        merged = simple_pivot.merge(
            rt_df[['IET #', 'RT']], 
            on='IET #', 
            how='left'
        )
        merged['RT'] = merged['RT'].fillna(0).astype(int)
        merged['DIFF'] = merged['SIMPLE'] - merged['RT']
        
        # Separate reconciled from discrepancies
        reconciled = merged[merged['DIFF'] == 0].copy()
        discrepancies = merged[merged['DIFF'] != 0].copy()
        
        # Sort discrepancies by absolute diff
        discrepancies['ABS_DIFF'] = discrepancies['DIFF'].abs()
        discrepancies = discrepancies.sort_values('ABS_DIFF', ascending=False)
        discrepancies = discrepancies.drop('ABS_DIFF', axis=1)
        
        # Not in RT
        not_in_rt = discrepancies[discrepancies['RT'] == 0].copy()
        
        # Totals
        total_simple = merged['SIMPLE'].sum()
        total_rt = merged['RT'].sum()
        total_diff = merged['DIFF'].sum()
        
        # Output file - same location as simple file
        output_dir = os.path.dirname(simple_file)
        output_file = os.path.join(output_dir, f"Reconciled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        # Write to Excel with all tabs
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Summary tab first
            summary_data = {
                'Metric': [
                    'Total SKUs',
                    'Reconciled (Match)',
                    'Discrepancies',
                    'Not in RT',
                    '',
                    'SIMPLE Total Qty',
                    'RT Total Qty', 
                    'Total Variance',
                    '',
                    'Reconciliation Date'
                ],
                'Value': [
                    len(merged),
                    len(reconciled),
                    len(discrepancies),
                    len(not_in_rt),
                    '',
                    total_simple,
                    total_rt,
                    total_diff,
                    '',
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Discrepancies tab
            if not discrepancies.empty:
                discrepancies.to_excel(writer, sheet_name='Discrepancies', index=False)
            
            # Reconciled tab
            if not reconciled.empty:
                reconciled.to_excel(writer, sheet_name='Reconciled', index=False)
            else:
                pd.DataFrame(columns=['IET #', 'SIMPLE', 'RT', 'DIFF']).to_excel(
                    writer, sheet_name='Reconciled', index=False
                )
            
            # Not in RT tab
            if not not_in_rt.empty:
                not_in_rt.to_excel(writer, sheet_name='Not in RT', index=False)
            
            # Original IE Tire detail data
            detail_df.to_excel(writer, sheet_name='IE Tire Detail', index=False)
            
            # Full comparison
            merged.to_excel(writer, sheet_name='Full Comparison', index=False)
        
        # Apply formatting
        self.format_output(output_file)
        
        return output_file
    
    def format_output(self, file_path):
        """Apply conditional formatting to the output file."""
        wb = load_workbook(file_path)
        
        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')
            
            max_width = 20 if sheet_name == 'IE Tire Detail' else 50
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws.column_dimensions[column_letter].width = min(max_length + 2, max_width)
            
            if sheet_name in ['Discrepancies', 'Reconciled', 'Not in RT', 'Full Comparison']:
                diff_col = None
                for idx, cell in enumerate(ws[1], 1):
                    if cell.value == 'DIFF':
                        diff_col = idx
                        break
                
                if diff_col:
                    for row in ws.iter_rows(min_row=2):
                        diff_cell = row[diff_col - 1]
                        try:
                            diff_val = float(diff_cell.value) if diff_cell.value else 0
                            if diff_val == 0:
                                for cell in row:
                                    cell.fill = green_fill
                            elif diff_val > 0:
                                for cell in row:
                                    cell.fill = yellow_fill
                            else:
                                for cell in row:
                                    cell.fill = red_fill
                        except:
                            pass
        
        wb.save(file_path)


if __name__ == '__main__':
    root = tk.Tk()
    app = ReconcilerApp(root)
    root.mainloop()
