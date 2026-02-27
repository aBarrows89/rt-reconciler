import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
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
        title = tk.Label(self.root, text="RT vs Simple Reconciler", font=("Arial", 18, "bold"), bg='#f0f0f0')
        title.pack(pady=20)

        frame1 = tk.Frame(self.root, bg='#f0f0f0', pady=10)
        frame1.pack(fill="x", padx=20)
        tk.Label(frame1, text="Simple Workbook:", font=("Arial", 10), bg='#f0f0f0').pack(anchor="w")
        f1 = tk.Frame(frame1, bg='#f0f0f0')
        f1.pack(fill="x", pady=5)
        tk.Entry(f1, textvariable=self.simple_file, width=55, font=("Arial", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="Browse...", command=self.browse_simple, width=10).pack(side="right", padx=5)

        frame2 = tk.Frame(self.root, bg='#f0f0f0', pady=10)
        frame2.pack(fill="x", padx=20)
        tk.Label(frame2, text="RT Scan Export:", font=("Arial", 10), bg='#f0f0f0').pack(anchor="w")
        f2 = tk.Frame(frame2, bg='#f0f0f0')
        f2.pack(fill="x", pady=5)
        tk.Entry(f2, textvariable=self.rt_file, width=55, font=("Arial", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f2, text="Browse...", command=self.browse_rt, width=10).pack(side="right", padx=5)

        self.status_var = tk.StringVar(value="Select files and click Reconcile")
        tk.Label(self.root, textvariable=self.status_var, font=("Arial", 10), bg='#f0f0f0').pack(pady=15)

        self.btn = tk.Button(
            self.root, text="RECONCILE", font=("Arial", 14, "bold"),
            width=20, height=2, bg='#4472C4', fg='white',
            activebackground='#3461b3', activeforeground='white',
            cursor='hand2', command=self.start_reconcile
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
        msg = (
            f"Reconciliation complete!\n\n"
            f"Original IE Tire rows: {stats['original']}\n"
            f"Previously Received: {stats['prev_received']}\n"
            f"Ready to Receive: {stats['ready']}\n"
            f"Unmatched Scans: {stats['unmatched']}\n"
            f"Remaining in IE Tire: {stats['remaining']}\n\n"
            f"Output: {os.path.basename(output_file)}\n\n"
            f"Open now?"
        )
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
        # --- Load IE Tire (full manifest) ---
        ie_df = pd.read_excel(simple_file, sheet_name='IE Tire')
        original_count = len(ie_df)

        # Clean IET # and part_number for matching
        ie_df['_clean_iet'] = ie_df['IET #'].apply(self.clean_part)
        if 'part_number' in ie_df.columns:
            ie_df['_clean_pn'] = ie_df['part_number'].apply(self.clean_part)
        else:
            ie_df['_clean_pn'] = ''

        # --- Load RT scan export and aggregate by part ---
        rt_raw = pd.read_excel(rt_file, sheet_name=0)
        part_col = None
        qty_col = None
        for c in rt_raw.columns:
            cl = str(c).lower().strip()
            if cl == 'part':
                part_col = c
            elif 'qty' in cl or 'quantity' in cl:
                qty_col = c
        if part_col is None:
            raise ValueError(f"Cannot find 'Part' column in RT file. Columns: {rt_raw.columns.tolist()}")
        if qty_col is None:
            qty_col = '_qty'
            rt_raw[qty_col] = 1
        else:
            rt_raw[qty_col] = pd.to_numeric(rt_raw[qty_col], errors='coerce').fillna(1).astype(int)

        rt_raw['_clean_part'] = rt_raw[part_col].apply(self.clean_part)
        rt_agg = rt_raw.groupby('_clean_part')[qty_col].sum().reset_index()
        rt_agg.columns = ['Part', 'RT_Qty']
        rt_agg = rt_agg[rt_agg['Part'] != '']

        # --- Build lookup: which IE Tire rows match each RT part ---
        # For each RT part, find matching IE Tire rows by IET # or part_number
        # Track which IE Tire rows get claimed
        ie_df['_matched'] = False
        ie_df['_match_status'] = ''  # 'full' or 'partial'

        unmatched_scans = []  # RT parts with no match in Simple at all
        rt_excess = []        # RT scanned more than Simple has

        for _, rt_row in rt_agg.iterrows():
            rt_part = rt_row['Part']
            rt_qty = int(rt_row['RT_Qty'])

            # Find matching IE Tire rows (by IET # first, then part_number)
            mask_iet = (ie_df['_clean_iet'] == rt_part) & (~ie_df['_matched'])
            mask_pn = (ie_df['_clean_pn'] == rt_part) & (~ie_df['_matched'])

            matching_idx = ie_df[mask_iet].index.tolist()
            if not matching_idx:
                matching_idx = ie_df[mask_pn].index.tolist()

            if not matching_idx:
                # No match in Simple at all -> Unmatched
                unmatched_scans.append({'Part': rt_part + '[', 'Qty': rt_qty})
                continue

            # Sum up Simple qty for these rows
            simple_qty = ie_df.loc[matching_idx, 'return_qty'].sum()
            simple_qty = int(simple_qty) if pd.notna(simple_qty) else len(matching_idx)

            if rt_qty >= simple_qty:
                # RT scanned all (or more than) what Simple has
                # All matching rows -> Previously Received
                ie_df.loc[matching_idx, '_matched'] = True
                ie_df.loc[matching_idx, '_match_status'] = 'full'
                # If RT has excess, track it as unmatched
                excess = rt_qty - simple_qty
                if excess > 0:
                    unmatched_scans.append({'Part': rt_part + '[', 'Qty': excess})
            else:
                # RT scanned fewer than Simple has (partial match)
                # Claim rows up to RT qty -> Ready to Receive
                claimed = 0
                for idx in matching_idx:
                    if claimed >= rt_qty:
                        break
                    row_qty = ie_df.at[idx, 'return_qty']
                    row_qty = int(row_qty) if pd.notna(row_qty) else 1
                    ie_df.at[idx, '_matched'] = True
                    ie_df.at[idx, '_match_status'] = 'partial'
                    claimed += row_qty

        # --- Split IE Tire into tabs ---
        prev_received_df = ie_df[ie_df['_match_status'] == 'full'].copy()
        ready_df = ie_df[ie_df['_match_status'] == 'partial'].copy()
        remaining_df = ie_df[~ie_df['_matched']].copy()
        unmatched_df = pd.DataFrame(unmatched_scans)

        # Drop internal columns before output
        drop_cols = ['_clean_iet', '_clean_pn', '_matched', '_match_status']
        prev_received_df = prev_received_df.drop(columns=drop_cols, errors='ignore')
        ready_df = ready_df.drop(columns=drop_cols, errors='ignore')
        remaining_df = remaining_df.drop(columns=drop_cols, errors='ignore')

        # Stats
        stats = {
            'original': original_count,
            'prev_received': len(prev_received_df),
            'ready': len(ready_df),
            'unmatched': len(unmatched_df),
            'remaining': len(remaining_df),
        }

        # --- Write output ---
        output_file = os.path.join(
            os.path.dirname(simple_file),
            f"Reconciled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            remaining_df.to_excel(writer, sheet_name='IE Tire', index=False)
            ready_df.to_excel(writer, sheet_name='Ready to Receive', index=False)
            unmatched_df.to_excel(writer, sheet_name='Unmatched', index=False)
            prev_received_df.to_excel(writer, sheet_name='Previously Received', index=False)

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
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)

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
