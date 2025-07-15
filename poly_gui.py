import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, Entry, IntVar, Checkbutton
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
        self.output_folder = None

        self.create_widgets()
        self.load_templates()

    def create_widgets(self):
        tk.Label(self.root, text="Poly", font=("Segoe UI", 20, "bold"), bg="#121212", fg="white").pack(pady=10)

        main_frame = tk.Frame(self.root, bg="#121212")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        template_outer = tk.Frame(main_frame, bg="#1e1e1e", width=350)
        template_outer.pack(side="left", fill="y", padx=(0, 10), pady=10)
        template_outer.pack_propagate(False)

        template_canvas = tk.Canvas(template_outer, bg="#1e1e1e", highlightthickness=0)
        template_scrollbar = tk.Scrollbar(template_outer, orient="vertical", command=template_canvas.yview)
        template_canvas.configure(yscrollcommand=template_scrollbar.set)
        template_scrollbar.pack(side="right", fill="y")
        template_canvas.pack(side="left", fill="both", expand=True)

        self.template_frame = tk.Frame(template_canvas, bg="#1e1e1e")
        self.template_frame.bind("<Configure>", lambda e: template_canvas.configure(scrollregion=template_canvas.bbox("all")))
        template_canvas.create_window((0, 0), window=self.template_frame, anchor="nw")

        field_outer = tk.Frame(main_frame, bg="#1e1e1e")
        field_outer.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)

        field_canvas = tk.Canvas(field_outer, bg="#1e1e1e", highlightthickness=0)
        field_scrollbar = tk.Scrollbar(field_outer, orient="vertical", command=field_canvas.yview)
        self.fields_frame = tk.Frame(field_canvas, bg="#1e1e1e")

        self.fields_frame.bind("<Configure>", lambda e: field_canvas.configure(scrollregion=field_canvas.bbox("all")))
        field_canvas.create_window((0, 0), window=self.fields_frame, anchor="nw")
        field_canvas.configure(yscrollcommand=field_scrollbar.set)
        field_canvas.pack(side="left", fill="both", expand=True)
        field_scrollbar.pack(side="right", fill="y")

        self.buttons_frame = tk.Frame(self.root, bg="#121212")
        self.buttons_frame.pack(pady=10)

        tk.Button(self.buttons_frame, text="Upload Template", command=self.upload_template,
                  bg="#2d2d2d", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#444").grid(row=0, column=0, padx=10)

        tk.Button(self.buttons_frame, text="Generate Document(s)", command=self.generate_documents,
                  bg="#0078d4", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#005ea2").grid(row=0, column=1, padx=10)

        tk.Button(self.buttons_frame, text="Sync Templates", command=self.sync_templates,
                  bg="#444444", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#666").grid(row=0, column=2, padx=10)

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

    def sync_templates(self):
        sync_folder = filedialog.askdirectory(title="Select Folder to Sync Templates From")
        if not sync_folder:
            return

        for file in os.listdir(sync_folder):
            if file.endswith(".docx"):
                source_path = os.path.join(sync_folder, file)
                dest_path = os.path.join(TEMPLATE_FOLDER, file)
                with open(source_path, "rb") as src, open(dest_path, "wb") as dst:
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
        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir:
            return

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

        os.startfile(output_dir)
        messagebox.showinfo("Success", f"Documents saved to '{output_dir}'.")

    def replace_placeholders(self, doc, replacements):
        pattern = re.compile(r"\{poly\.([a-zA-Z0-9_ ]+)\}")

        def split_runs_by_char(paragraph):
            chars = []
            for run in paragraph.runs:
                for c in run.text:
                    chars.append({
                        "char": c,
                        "bold": run.bold,
                        "italic": run.italic,
                        "underline": run.underline
                    })
            return chars

        def apply_replacement(paragraph):
            chars = split_runs_by_char(paragraph)
            full_text = ''.join(c['char'] for c in chars)

            matches = list(pattern.finditer(full_text))
            if not matches:
                return

            new_chars = []
            last_idx = 0

            for match in matches:
                start, end = match.span()
                field = match.group(1)
                replacement = replacements.get(field, match.group(0))

                new_chars.extend(chars[last_idx:start])
                fmt = chars[start] if start < len(chars) else {"bold": None, "italic": None, "underline": None}
                new_chars.extend([{
                    "char": rc,
                    "bold": fmt['bold'],
                    "italic": fmt['italic'],
                    "underline": fmt['underline']
                } for rc in replacement])
                last_idx = end

            new_chars.extend(chars[last_idx:])

            for run in paragraph.runs:
                run.text = ""
            paragraph._element.clear_content()

            new_run = None
            for ch in new_chars:
                if (
                    new_run is None or
                    new_run.bold != ch['bold'] or
                    new_run.italic != ch['italic'] or
                    new_run.underline != ch['underline']
                ):
                    new_run = paragraph.add_run()
                    new_run.bold = ch['bold']
                    new_run.italic = ch['italic']
                    new_run.underline = ch['underline']
                new_run.text += ch['char']

        for para in doc.paragraphs:
            apply_replacement(para)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        apply_replacement(para)

if __name__ == "__main__":
    root = tk.Tk()
    app = PolyApp(root)
    root.mainloop()
