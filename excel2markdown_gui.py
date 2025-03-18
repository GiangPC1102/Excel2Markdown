#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import sys
import os
from pathlib import Path

class Excel2MarkdownGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel2Markdown Converter")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Variables
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.include_index_var = tk.BooleanVar(value=False)
        self.selected_sheet_var = tk.StringVar()
        self.sheets = []
        
        # Create UI
        self.create_widgets()
    
    def create_widgets(self):
        # Main layout frames
        input_frame = ttk.LabelFrame(self.root, text="Input/Output")
        input_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        options_frame = ttk.LabelFrame(self.root, text="Options")
        options_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        preview_frame = ttk.LabelFrame(self.root, text="Markdown Preview")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", expand=False, padx=10, pady=10)
        
        # Input file selection
        ttk.Label(input_frame, text="Excel File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Output file selection
        ttk.Label(input_frame, text="Output File:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.output_file_var, width=50).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(input_frame, text="Browse...", command=self.browse_output_file).grid(row=1, column=2, padx=5, pady=5)
        
        input_frame.columnconfigure(1, weight=1)
        
        # Options
        ttk.Label(options_frame, text="Sheet:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.sheet_combobox = ttk.Combobox(options_frame, textvariable=self.selected_sheet_var, state="readonly")
        self.sheet_combobox.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Checkbutton(options_frame, text="Include Index", variable=self.include_index_var).grid(row=0, column=2, sticky="w", padx=20, pady=5)
        
        options_frame.columnconfigure(1, weight=1)
        
        # Preview
        self.preview_text = ScrolledText(preview_frame, wrap=tk.WORD)
        self.preview_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Buttons
        ttk.Button(button_frame, text="Preview", command=self.preview_markdown).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Convert", command=self.convert_to_markdown).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.destroy).pack(side="right", padx=5)
    
    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.input_file_var.set(file_path)
            self.load_sheets(file_path)
            
            # Auto-generate output file name
            input_path = Path(file_path)
            output_path = input_path.with_suffix('.md')
            self.output_file_var.set(str(output_path))
    
    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Markdown File",
            filetypes=[("Markdown files", "*.md"), ("Text files", "*.txt"), ("All files", "*.*")],
            defaultextension=".md"
        )
        if file_path:
            self.output_file_var.set(file_path)
    
    def load_sheets(self, excel_file):
        try:
            xls = pd.ExcelFile(excel_file)
            self.sheets = xls.sheet_names
            self.sheet_combobox['values'] = ['All Sheets'] + self.sheets
            self.sheet_combobox.current(0)  # Select 'All Sheets' by default
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")
            self.sheets = []
            self.sheet_combobox['values'] = []
    
    def get_selected_sheet(self):
        selected = self.selected_sheet_var.get()
        if selected == 'All Sheets':
            return None  # None means all sheets
        return selected
    
    def preview_markdown(self):
        excel_file = self.input_file_var.get()
        if not excel_file:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        try:
            sheet_name = self.get_selected_sheet()
            include_index = self.include_index_var.get()
            
            # Get all sheet names if not specified
            if sheet_name is None:
                sheet_names = self.sheets
            else:
                sheet_names = [sheet_name]
            
            # Initialize output string
            md_content = ""
            
            # Process each sheet
            for sheet in sheet_names:
                # Read the Excel file
                df = pd.read_excel(excel_file, sheet_name=sheet)
                
                # Add sheet name as header if multiple sheets
                if len(sheet_names) > 1:
                    md_content += f"## {sheet}\n\n"
                
                # Convert dataframe to markdown
                md_table = df.to_markdown(index=include_index)
                md_content += md_table + "\n\n"
            
            # Display in preview
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, md_content)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate preview: {e}")
    
    def convert_to_markdown(self):
        excel_file = self.input_file_var.get()
        output_file = self.output_file_var.get()
        
        if not excel_file:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        if not output_file:
            messagebox.showwarning("Warning", "Please specify an output file.")
            return
        
        try:
            sheet_name = self.get_selected_sheet()
            include_index = self.include_index_var.get()
            
            # Get all sheet names if not specified
            if sheet_name is None:
                sheet_names = self.sheets
            else:
                sheet_names = [sheet_name]
            
            # Initialize output string
            md_content = ""
            
            # Process each sheet
            for sheet in sheet_names:
                # Read the Excel file
                df = pd.read_excel(excel_file, sheet_name=sheet)
                
                # Add sheet name as header if multiple sheets
                if len(sheet_names) > 1:
                    md_content += f"## {sheet}\n\n"
                
                # Convert dataframe to markdown
                md_table = df.to_markdown(index=include_index)
                md_content += md_table + "\n\n"
            
            # Write to file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(md_content)
            
            messagebox.showinfo("Success", f"Excel file converted to Markdown: {output_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert file: {e}")

def main():
    root = tk.Tk()
    app = Excel2MarkdownGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 