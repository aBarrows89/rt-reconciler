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
        se
