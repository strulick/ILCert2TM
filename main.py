import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from datetime import datetime
import csv  # for saving

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TM SOM Extractor")
        self.geometry("600x450")

        # Email file chooser
        tk.Label(self, text="Email File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.eml_var = tk.StringVar()
        tk.Entry(self, textvariable=self.eml_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self, text="Browse", command=self.browse_eml).grid(row=0, column=2, padx=5, pady=5)

        # Description prefix
        tk.Label(self, text="Desc Prefix:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.desc_var = tk.StringVar()
        tk.Entry(self, textvariable=self.desc_var, width=50).grid(row=1, column=1, columnspan=2, padx=5, pady=5)

        # Output folder chooser
        tk.Label(self, text="Output Folder:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.out_var = tk.StringVar(value=os.getcwd())
        tk.Entry(self, textvariable=self.out_var, width=50).grid(row=2, column=1, padx=5, pady=5)
        tk.Button(self, text="Browse", command=self.browse_out).grid(row=2, column=2, padx=5, pady=5)

        # Run button
        tk.Button(self, text="Run", width=10, command=self.run).grid(row=3, column=1, pady=10)

        # Console output
        self.console = scrolledtext.ScrolledText(self, wrap=tk.WORD, height=18)
        self.console.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.grid_rowconfigure(4, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def browse_eml(self):
        path = filedialog.askopenfilename(filetypes=[("Email files", ("*.eml", "*.msg"))])
        if path:
            self.eml_var.set(path)

    def browse_out(self):
        path = filedialog.askdirectory()
        if path:
            self.out_var.set(path)

    def run(self):
        eml_path = self.eml_var.get()
        out_dir = self.out_var.get()
        desc = self.desc_var.get() or None

        if not eml_path:
            messagebox.showwarning("Missing File", "Please select an .eml or .msg file.")
            return
        if not os.path.isdir(out_dir):
            messagebox.showwarning("Invalid Folder", "Please select a valid folder.")
            return

        self.console.insert(tk.END, "Loading modules…\n")
        self.update_idletasks()

        try:
            import prepTM
            self.console.insert(tk.END, f"Processing: {eml_path}\n")
            entries = prepTM.extract_som_entries(eml_path, desc)

            ts = datetime.now().strftime("%Y%m%d%H%M%S")
            filename = f"IOC{ts}.csv"
            out_path = os.path.join(out_dir, filename)

            with open(out_path, 'w', newline='') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=["Type","Object","Description"])
                writer.writeheader()
                writer.writerows(entries)

            self.console.insert(tk.END, f"✅ Saved {len(entries)} entries to {out_path}\n")
        except Exception as e:
            self.console.insert(tk.END, f"❌ Error: {e}\n")

if __name__ == "__main__":
    App().mainloop()
