import os
import re
import sys
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
        self.root.configure(bg="#0f0f0f")
        tk.Label(self.root, text="Poly", font=("Segoe UI", 24, "bold"), bg="#0f0f0f", fg="white").pack(pady=(20, 10))

        main_frame = tk.Frame(self.root, bg="#0f0f0f")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        left_panel = tk.Frame(main_frame, bg="#1a1a1a", width=370, height=600, bd=0, relief="flat")
        left_panel.grid(row=0, column=0, sticky="ns")
        left_panel.grid_propagate(False)

        self.template_canvas = tk.Canvas(left_panel, bg="#1a1a1a", highlightthickness=0, bd=0)
        template_scrollbar = tk.Scrollbar(left_panel, orient="vertical", command=self.template_canvas.yview, bg="#2a2a2a")
        self.template_canvas.configure(yscrollcommand=template_scrollbar.set)

        self.template_inner_frame = tk.Frame(self.template_canvas, bg="#1a1a1a")
        self.template_canvas.create_window((0, 0), window=self.template_inner_frame, anchor="nw")
        self.template_inner_frame.bind("<Configure>", lambda e: self.template_canvas.configure(scrollregion=self.template_canvas.bbox("all")))

        self.template_canvas.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        template_scrollbar.pack(side="right", fill="y")

        right_panel = tk.Frame(main_frame, bg="#1a1a1a", bd=0)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(20, 0))
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)

        self.fields_canvas = tk.Canvas(right_panel, bg="#1a1a1a", highlightthickness=0, bd=0)
        fields_scrollbar = tk.Scrollbar(right_panel, orient="vertical", command=self.fields_canvas.yview, bg="#2a2a2a")
        self.fields_canvas.configure(yscrollcommand=fields_scrollbar.set)

        self.fields_inner_frame = tk.Frame(self.fields_canvas, bg="#1a1a1a")
        self.fields_canvas.create_window((0, 0), window=self.fields_inner_frame, anchor="nw")
        self.fields_inner_frame.bind("<Configure>", lambda e: self.fields_canvas.configure(scrollregion=self.fields_canvas.bbox("all")))

        self.fields_canvas.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        fields_scrollbar.pack(side="right", fill="y")

        self.buttons_frame = tk.Frame(self.root, bg="#0f0f0f")
        self.buttons_frame.pack(pady=(0, 25))

        style = {
            "font": ("Segoe UI", 10, "bold"),
            "width": 20,
            "height": 2,
            "relief": "flat",
            "bd": 0,
            "cursor": "hand2"
        }

        tk.Button(self.buttons_frame, text="Upload Template", command=self.upload_template,
                bg="#2a2a2a", fg="white", activebackground="#3c3c3c", activeforeground="white", **style
        ).grid(row=0, column=0, padx=15)

        tk.Button(self.buttons_frame, text="Generate Document(s)", command=self.generate_documents,
                bg="#d32f2f", fg="white", activebackground="#b71c1c", activeforeground="white", **style
        ).grid(row=0, column=1, padx=15)



    def load_templates(self):
        for widget in self.template_inner_frame.winfo_children():
            widget.destroy()

        if not os.path.exists(TEMPLATE_FOLDER):
            os.makedirs(TEMPLATE_FOLDER)

        files = [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith(".docx")]
        for file in sorted(files):
            var = IntVar()
            cb = Checkbutton(self.template_inner_frame, text=file, variable=var,
                             font=("Segoe UI", 10), anchor="w", command=self.update_selected_templates,
                             bg="#1e1e1e", fg="white", wraplength=330, justify="left",
                             selectcolor="#1e1e1e", activebackground="#1e1e1e")
            cb.var = var
            cb.filename = file
            cb.pack(fill="x", anchor="w", padx=10, pady=2)

    def update_selected_templates(self):
        self.selected_templates = []
        for widget in self.template_inner_frame.winfo_children():
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
        for widget in self.fields_inner_frame.winfo_children():
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
            tk.Label(self.fields_inner_frame, text=field.replace("_", " "), bg="#1e1e1e", fg="white", font=("Segoe UI", 10)).pack(anchor="w", padx=10, pady=(10, 2))
            entry = Entry(self.fields_inner_frame, width=50, bg="#2a2a2a", fg="white", insertbackground="white", font=("Segoe UI", 10))
            entry.pack(anchor="w", padx=10, pady=2)
            self.entries[field] = entry

    def generate_documents(self):
        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir:
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
        try:
            if os.name == 'nt':
                os.startfile(output_dir)
            elif os.name == 'posix':
                import subprocess
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', output_dir])
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not open folder:\n{str(e)}")

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
