import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime
import os

class ReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RT Reconciler")
        self.root.geometry("520x400")
        self.root.resizable(False, False)
        self.root.configure(bg='#f0f0f0')
        self.simple_file = tk.StringVar()
        self.rt_file = tk.StringVar()
        self.create_widgets()
    
    def create_widgets(self):
        # Title
        title = tk.Label(self.root, text="RT vs Simple Reconciler", font=("Arial", 18, "bold"), bg='#f0f0f0')
        title.pack(pady=20)
        
        # Simple Workbook
        frame1 = tk.Frame(self.root, bg='#f0f0f0', pady=10)
        frame1.pack(fill="x", padx=20)
        tk.Label(frame1, text="Simple Workbook:", font=("Arial", 10), bg='#f0f0f0').pack(anchor="w")
        f1 = tk.Frame(frame1, bg='#f0f0f0')
        f1.pack(fill="x", pady=5)
        tk.Entry(f1, textvariable=self.simple_file, width=55, font=("Arial", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="Browse...", command=self.browse_simple, width=10).pack(side="right", padx=5)
        
        # RT Export
        frame2 = tk.Frame(self.root, bg='#f0f0f0', pady=10)
        frame2.pack(fill="x", padx=20)
        tk.Label(frame2, text="RT Scan Export:", font=("Arial", 10), bg='#f0f0f0').pack(anchor="w")
        f2 = tk.Frame(frame2, bg='#f0f0f0')
        f2.pack(fill="x", pady=5)
        tk.Entry(f2, textvariable=self.rt_file, width=55, font=("Arial", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f2, text="Browse...", command=self.browse_rt, width=10).pack(side="right", padx=5)
        
        # Status
        self.status_var = tk.StringVar(value="Select files and click Reconcile")
        tk.Label(self.root, textvariable=self.status_var, font=("Arial", 10), bg='#f0f0f0').pack(pady=15)
        
        # Big visible button
        self.btn = tk.Button(
            self.root, 
            text="RECONCILE", 
            font=("Arial", 14, "bold"),
            width=20,
            height=2,
            bg='#4472C4',
            fg='white',
            activebackground='#3461b3',
            activeforeground='white',
            cursor='hand2',
            command=self.start_reconcile
        )
        self.btn.pack(pady=20)
    
    def browse_simple(self):
        f = filedialog.askopenfilename(title="Select Simple Workbook", filetypes=[("Excel", "*.xlsx *.xls")])
        if f: self.simple_file.set(f)
    
    def browse_rt(self):
        f = filedialog.askopenfilename(title="Select RT Scan Export", filetypes=[("Excel", "*.xlsx *.xls")])
        if f: self.rt_file.set(f)
    
    def start_reconcile(self):
        if not self.simple_file.get() or not self.rt_file.get():
            messagebox.showerror("Error", "Please select both files")
            return
        self.btn.config(state='disabled')
        self.status_var.set("Processing...")
        self.root.update()
        try:
            output, stats = self.reconcile(self.simple_file.get(), self.rt_file.get())
            self.on_complete(output, stats)
        except Exception as e:
            import traceback
            self.on_error(f"{str(e)}\n\n{traceback.format_exc()}")
    
    def on_complete(self, output_file, stats):
        self.btn.config(state='normal')
        self.status_var.set("Complete!")
        
        msg = f"""Reconciliation complete!

RT Scans: {stats['rt_scans']}
Matched: {stats['matched']}
Ready to Receive: {stats['ready']}
Unmatched (in RT, not Simple): {stats['unmatched']}
Previously Received: {stats['received']}

Output: {os.path.basename(output_file)}

Open now?"""
        
        if messagebox.askyesno("Success", msg):
            os.startfile(output_file)
    
    def on_error(self, msg):
        self.btn.config(state='normal')
        self.status_var.set("Error")
        messagebox.showerror("Error", msg)
    
    @staticmethod
    def clean_part(val):
        """Clean part numbers: strip whitespace, trailing [, uppercase."""
        if pd.isna(val):
            return ''
        return str(val).strip().rstrip('[').upper()
    
    def reconcile(self, simple_file, rt_file):
        # Load Simple Workbook - IE Tire sheet has the manifest
        detail_df = pd.read_excel(simple_file, sheet_name='IE Tire')
        
        # Load RT scan export (columns: Start DT, Part, Part Qty)
        rt_df = pd.read_excel(rt_file, sheet_name=0)
        
        # Find the part column in RT (could be 'Part', 'part', etc.)
        part_col = None
        qty_col = None
        for c in rt_df.columns:
            cl = str(c).lower().strip()
            if cl == 'part':
                part_col = c
            elif 'qty' in cl or 'quantity' in cl:
                qty_col = c
        
        if part_col is None:
            raise ValueError(f"Cannot find 'Part' column in RT file. Columns found: {rt_df.columns.tolist()}")
        
        # Clean RT parts and aggregate by part number
        rt_df['_clean_part'] = rt_df[part_col].apply(self.clean_part)
        if qty_col:
            rt_df[qty_col] = pd.to_numeric(rt_df[qty_col], errors='coerce').fillna(1).astype(int)
        else:
            qty_col = '_qty'
            rt_df[qty_col] = 1
        
        rt_agg = rt_df.groupby('_clean_part')[qty_col].sum().reset_index()
        rt_agg.columns = ['Part', 'RT_Qty']
        rt_agg = rt_agg[rt_agg['Part'] != '']
        
        # Build lookup sets from Simple: IET # and part_number
        detail_df['_clean_iet'] = detail_df['IET #'].apply(self.clean_part)
        if 'part_number' in detail_df.columns:
            detail_df['_clean_pn'] = detail_df['part_number'].apply(self.clean_part)
        else:
            detail_df['_clean_pn'] = ''
        
        # Aggregate Simple quantities by cleaned IET #
        # First build a combined lookup: IET # -> total qty, part_number -> total qty
        simple_by_iet = detail_df.groupby('_clean_iet')['return_qty'].sum().to_dict()
        simple_by_pn = detail_df.groupby('_clean_pn')['return_qty'].sum().to_dict()
        
        # All known Simple part identifiers
        simple_all_parts = set(simple_by_iet.keys()) | set(simple_by_pn.keys())
        simple_all_parts.discard('')
        
        # Reconcile each RT part
        results = []
        for _, row in rt_agg.iterrows():
            part = row['Part']
            rt_qty = int(row['RT_Qty'])
            
            # Try matching IET # first, then part_number
            simple_qty = 0
            matched_via = ''
            if part in simple_by_iet:
                simple_qty = int(simple_by_iet[part])
                matched_via = 'IET #'
            elif part in simple_by_pn:
                simple_qty = int(simple_by_pn[part])
                matched_via = 'part_number'
            
            diff = simple_qty - rt_qty
            
            if simple_qty == 0:
                status = 'Unmatched'
            elif diff == 0:
                status = 'Previously Received'
            else:
                status = 'Ready to Receive'
            
            results.append({
                'Part': part,
                'Simple_Qty': simple_qty,
                'RT_Qty': rt_qty,
                'DIFF': diff,
                'Status': status,
                'Matched Via': matched_via
            })
        
        comparison_df = pd.DataFrame(results)
        
        # Split into tabs
        ready_df = comparison_df[comparison_df['Status'] == 'Ready to Receive'].copy()
        unmatched_df = comparison_df[comparison_df['Status'] == 'Unmatched'].copy()
        received_df = comparison_df[comparison_df['Status'] == 'Previously Received'].copy()
        
        # Stats
        stats = {
            'rt_scans': int(rt_df[qty_col].sum()),
            'matched': int(comparison_df[comparison_df['Status'] != 'Unmatched']['RT_Qty'].sum()),
            'ready': len(ready_df),
            'unmatched': len(unmatched_df),
            'received': len(received_df),
        }
        
        # Output file
        output_file = os.path.join(
            os.path.dirname(simple_file),
            f"Reconciled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # IE Tire - original detail data
            out_detail = detail_df.drop(columns=['_clean_iet', '_clean_pn'], errors='ignore')
            out_detail.to_excel(writer, sheet_name='IE Tire', index=False)
            
            # Ready to Receive - qty mismatch between Simple and RT
            ready_df.drop(columns=['Status'], errors='ignore').to_excel(
                writer, sheet_name='Ready to Receive', index=False)
            
            # Unmatched - in RT but not in Simple at all
            unmatched_df.drop(columns=['Status'], errors='ignore').to_excel(
                writer, sheet_name='Unmatched', index=False)
            
            # Previously Received - fully matched
            received_df.drop(columns=['Status'], errors='ignore').to_excel(
                writer, sheet_name='Previously Received', index=False)
        
        self.format_workbook(output_file)
        
        return output_file, stats
    
    def format_workbook(self, file_path):
        wb = load_workbook(file_path)
        
        header_fill = PatternFill('solid', fgColor='4472C4')
        header_font = Font(color='FFFFFF', bold=True)
        red = PatternFill('solid', fgColor='FFC7CE')
        yellow = PatternFill('solid', fgColor='FFEB9C')
        green = PatternFill('solid', fgColor='C6EFCE')
        
        for ws in wb.worksheets:
            # Headers
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            # Column widths
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)
            
            # Color code by tab
            if ws.title == 'Ready to Receive':
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.fill = yellow
            
            elif ws.title == 'Unmatched':
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.fill = red
            
            elif ws.title == 'Previously Received':
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        cell.fill = green
        
        wb.save(file_path)

if __name__ == '__main__':
    root = tk.Tk()
    ReconcilerApp(root)
    root.mainloop()
