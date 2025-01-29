import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog

class ExcelSorter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sorter")
        
        self.file_path = ""
        self.df = None
        self.sort_order = {}
        
        self.load_button = tk.Button(root, text="Load Excel File", command=self.load_file)
        self.load_button.pack(pady=10)
        
        self.tree = ttk.Treeview(root)
        self.tree.pack(expand=True, fill='both')
        
    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            self.df = pd.read_excel(self.file_path)
            self.sort_order = {col: True for col in self.df.columns}  
            self.display_table()

    def display_table(self):
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = list(self.df.columns)
        self.tree['show'] = "headings"
        
        for col in self.df.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c))
            self.tree.column(col, width=100)
        
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))
        
    def sort_by(self, column):
        self.df.sort_values(by=column, ascending=self.sort_order[column], inplace=True)
        self.sort_order[column] = not self.sort_order[column] 
        self.display_table()

root = tk.Tk()
app = ExcelSorter(root)
root.mainloop()