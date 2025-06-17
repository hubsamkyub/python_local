import tkinter as tk
from tkinter import ttk


class ScrollableCheckList(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.items = {}

    def add_item(self, text, checked=False):
        var = tk.BooleanVar(value=checked)
        cb = ttk.Checkbutton(self.scrollable_frame, text=text, variable=var)
        cb.pack(anchor="w", fill="x")
        self.items[text] = var

    def get_checked_items(self):
        return [text for text, var in self.items.items() if var.get()]

    def clear(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.items = {}        
