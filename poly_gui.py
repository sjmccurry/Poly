"""
Poly Document Generator

Author: Spencer McCurry
Organization: WMM Legal
Created: 2025

Description:
This software was developed by Spencer McCurry for use by WMM Legal.
It is designed to manage Word document templates, detect and fill
custom variables using a GUI interface, and preserve formatting.

All rights reserved. Unauthorized copying, distribution, or modification
of this software is strictly prohibited without prior written consent
from the author or WMM Legal.

© 2025 Spencer McCurry / WMM Legal
"""




import os
import re
try:
    import tkinter as tk
except:
    os.system('pip3 install tkinter')
    import tkinter as tk
from tkinter import filedialog, messagebox, Entry, IntVar, Checkbutton
try:
    from docx import Document
except:
    os.system('pip3 install python-docx')
    from docx import Document

TEMPLATE_FOLDER = "templates"

class PolyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Poly Document Generator")
        self.root.geometry("1200x700")
        self.root.configure(bg="#121212")

        self.selected_templates = []
        self.entries = {}

        self.create_widgets()
        self.load_templates()

    def create_widgets(self):
        tk.Label(self.root, text="Poly Document Generator", font=("Segoe UI", 20, "bold"), bg="#121212", fg="white").pack(pady=10)

        main_frame = tk.Frame(self.root, bg="#121212")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.template_frame = tk.Frame(main_frame, bg="#1e1e1e", width=350)
        self.template_frame.pack(side="left", fill="y", padx=(0, 10), pady=10)
        self.template_frame.pack_propagate(False)

        self.fields_frame = tk.Frame(main_frame, bg="#1e1e1e")
        self.fields_frame.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)

        self.buttons_frame = tk.Frame(self.root, bg="#121212")
        self.buttons_frame.pack(pady=10)

        tk.Button(self.buttons_frame, text="Upload Template", command=self.upload_template,
                  bg="#2d2d2d", fg="white", font=("Segoe UI", 10, "bold"),
                  activebackground="#444").grid(row=0, column=0, padx=10)

        tk.Button(self.buttons_frame, text="Generate Document(s)", command=self.generate_documents,
                  bg="#0078d4", fg="white", font=("Segoe UI", 10, "bold"),
                  activebackground="#005ea2").grid(row=0, column=1, padx=10)

    def load_templates(self):
        for widget in self.template_frame.winfo_children():
            widget.destroy()

        if not os.path.exists(TEMPLATE_FOLDER):
            os.makedirs(TEMPLATE_FOLDER)

        files = [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith(".docx")]
        for file in files:
            var = IntVar()
            cb = Checkbutton(self.template_frame, text=file, variable=var,
                             font=("Segoe UI", 10), anchor="w", command=self.update_selected_templates,
                             bg="#1e1e1e", fg="white", wraplength=330, justify="left",
                             selectcolor="#1e1e1e", activebackground="#1e1e1e")
            cb.var = var
            cb.filename = file
            cb.pack(fill="x", anchor="w", padx=10, pady=2)

    def update_selected_templates(self):
        self.selected_templates = []
        for widget in self.template_frame.winfo_children():
            if isinstance(widget, Checkbutton) and widget.var.get() == 1:
                self.selected_templates.append(widget.filename)
        self.load_fields()

    def upload_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            filename = os.path.basename(file_path)
            dest_path = os.path.join(TEMPLATE_FOLDER, filename)
            with open(file_path, "rb") as src, open(dest_path, "wb") as dst:
                dst.write(src.read())
            self.load_templates()

    def extract_poly_fields(self, document):
        pattern = r"\{poly\.([a-zA-Z0-9_ ]+)\}"
        fields = set()
        for para in document.paragraphs:
            matches = re.findall(pattern, para.text)
            fields.update(matches)
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    matches = re.findall(pattern, cell.text)
                    fields.update(matches)
        return fields

    def load_fields(self):
        for widget in self.fields_frame.winfo_children():
            widget.destroy()

        all_fields = set()
        for filename in self.selected_templates:
            try:
                doc = Document(os.path.join(TEMPLATE_FOLDER, filename))
                fields = self.extract_poly_fields(doc)
                all_fields.update(fields)
            except Exception:
                continue

        self.entries = {}
        for field in sorted(all_fields):
            tk.Label(self.fields_frame, text=field.replace("_", " "), bg="#1e1e1e", fg="white", font=("Segoe UI", 10)).pack(anchor="w", padx=10, pady=(10, 2))
            entry = Entry(self.fields_frame, width=50, bg="#2a2a2a", fg="white", insertbackground="white", font=("Segoe UI", 10))
            entry.pack(anchor="w", padx=10, pady=2)
            self.entries[field] = entry

    def generate_documents(self):
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)

        values = {key: entry.get() for key, entry in self.entries.items()}

        for filename in self.selected_templates:
            try:
                doc_path = os.path.join(TEMPLATE_FOLDER, filename)
                doc = Document(doc_path)
                self.replace_placeholders(doc, values)
                new_name = filename.replace(".docx", "_filled.docx")
                output_path = os.path.join(output_dir, new_name)
                doc.save(output_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not process {filename}:\n{str(e)}")
                return

        messagebox.showinfo("Success", f"Documents saved to '{output_dir}' folder.")


    def replace_placeholders(self, doc, replacements):
        def replace_text_in_run(run):
            for key, value in replacements.items():
                placeholder = f"{{poly.{key}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)

        for para in doc.paragraphs:
            for run in para.runs:
                replace_text_in_run(run)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            replace_text_in_run(run)

if __name__ == "__main__":
    root = tk.Tk()
    app = PolyApp(root)
    root.mainloop()
