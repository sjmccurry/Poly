import os
import re
import json
import tkinter as tk
from tkinter import filedialog, messagebox, Entry, IntVar, Checkbutton
from docx import Document

CONFIG_FILE = "poly_config.json"
LOCAL_VERSION_FILE = "version.txt"
DEFAULT_TEMPLATE_FOLDER = "templates"
DEFAULT_OUTPUT_FOLDER = os.getcwd()

def read_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return "vUnknown"

class PolyApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Poly Document Generator {read_version()}")
        self.root.geometry("1280x720")
        self.root.configure(bg="#121212")

        self.selected_templates = []
        self.entries = {}
        self.template_folder = DEFAULT_TEMPLATE_FOLDER
        self.output_folder = DEFAULT_OUTPUT_FOLDER

        self.load_config()

        self.create_widgets()
        self.load_templates()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                    self.template_folder = data.get("template_folder", DEFAULT_TEMPLATE_FOLDER)
                    self.output_folder = data.get("output_folder", DEFAULT_OUTPUT_FOLDER)
            except:
                self.template_folder = DEFAULT_TEMPLATE_FOLDER
                self.output_folder = DEFAULT_OUTPUT_FOLDER
        else:
            self.template_folder = DEFAULT_TEMPLATE_FOLDER
            self.output_folder = DEFAULT_OUTPUT_FOLDER
            self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump({
                    "template_folder": self.template_folder,
                    "output_folder": self.output_folder
                }, f)
        except:
            pass

    def create_widgets(self):
        header = tk.Frame(self.root, bg="#1f1f1f")
        header.pack(fill="x")
        tk.Label(header, text="Poly Document Generator", font=("Segoe UI", 24, "bold"), bg="#1f1f1f", fg="white").pack(pady=10)

        main_frame = tk.Frame(self.root, bg="#121212")
        main_frame.pack(fill="both", expand=True, padx=15, pady=10)

        # Left: Templates
        left_frame = tk.Frame(main_frame, bg="#1e1e1e", width=400)
        left_frame.pack(side="left", fill="y", padx=(0, 10))
        left_frame.pack_propagate(False)

        tk.Label(left_frame, text="Templates", bg="#1e1e1e", fg="white", font=("Segoe UI", 12, "bold")).pack(pady=10)

        template_canvas = tk.Canvas(left_frame, bg="#1e1e1e", highlightthickness=0)
        template_scrollbar = tk.Scrollbar(left_frame, orient="vertical", command=template_canvas.yview)
        template_canvas.configure(yscrollcommand=template_scrollbar.set)
        template_scrollbar.pack(side="right", fill="y")
        template_canvas.pack(side="left", fill="both", expand=True)

        self.template_frame = tk.Frame(template_canvas, bg="#1e1e1e")
        self.template_frame.bind("<Configure>", lambda e: template_canvas.configure(scrollregion=template_canvas.bbox("all")))
        template_canvas.create_window((0, 0), window=self.template_frame, anchor="nw")

        # Center: Fields
        center_frame = tk.Frame(main_frame, bg="#1e1e1e")
        center_frame.pack(side="left", fill="both", expand=True)

        tk.Label(center_frame, text="Fill in Fields", bg="#1e1e1e", fg="white", font=("Segoe UI", 12, "bold")).pack(pady=10)

        field_canvas = tk.Canvas(center_frame, bg="#1e1e1e", highlightthickness=0)
        field_scrollbar = tk.Scrollbar(center_frame, orient="vertical", command=field_canvas.yview)
        self.fields_frame = tk.Frame(field_canvas, bg="#1e1e1e")

        self.fields_frame.bind("<Configure>", lambda e: field_canvas.configure(scrollregion=field_canvas.bbox("all")))
        field_canvas.create_window((0, 0), window=self.fields_frame, anchor="nw")
        field_canvas.configure(yscrollcommand=field_scrollbar.set)
        field_canvas.pack(side="left", fill="both", expand=True)
        field_scrollbar.pack(side="right", fill="y")

        # Right: Actions
        right_frame = tk.Frame(main_frame, bg="#1e1e1e", width=200)
        right_frame.pack(side="left", fill="y", padx=(10, 0))
        right_frame.pack_propagate(False)

        tk.Label(right_frame, text="Actions", bg="#1e1e1e", fg="white", font=("Segoe UI", 12, "bold")).pack(pady=10)

        button_style = {"bg": "#2d2d2d", "fg": "white", "font": ("Segoe UI", 10, "bold"), "activebackground": "#444", "width": 20, "anchor": "w"}
        highlight_style = button_style.copy()
        highlight_style["bg"] = "#0078d4"
        highlight_style["activebackground"] = "#005ea2"

        tk.Button(right_frame, text="Generate Document(s)", command=self.generate_documents, **highlight_style).pack(pady=10)
        tk.Button(right_frame, text="Upload Template", command=self.upload_template, **button_style).pack(pady=10)


        tk.Button(right_frame, text="Select Template Folder", command=self.select_template_folder, **button_style).pack(pady=10)

    def select_template_folder(self):
        folder = filedialog.askdirectory(title="Select Template Folder")
        if folder:
            self.template_folder = folder
            self.save_config()
            self.load_templates()

    def load_templates(self):
        for widget in self.template_frame.winfo_children():
            widget.destroy()

        if not os.path.exists(self.template_folder):
            os.makedirs(self.template_folder)

        files = [f for f in os.listdir(self.template_folder) if f.endswith(".docx")]
        for file in files:
            var = IntVar()
            cb = Checkbutton(self.template_frame, text=file, variable=var,
                             font=("Segoe UI", 10), anchor="w", command=self.update_selected_templates,
                             bg="#1e1e1e", fg="white", wraplength=280, justify="left",
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
            dest_path = os.path.join(self.template_folder, filename)
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
                doc = Document(os.path.join(self.template_folder, filename))
                fields = self.extract_poly_fields(doc)
                all_fields.update(fields)
            except Exception:
                continue

        self.entries = {}
        for field in sorted(all_fields):
            label_text = re.sub(r'([a-z])([A-Z])', r'\1 \2', field).replace("_", " ")
            tk.Label(self.fields_frame, text=label_text, bg="#1e1e1e", fg="white", font=("Segoe UI", 10)).pack(anchor="w", padx=10, pady=(10, 2))
            entry = Entry(self.fields_frame, width=50, bg="#2a2a2a", fg="white", insertbackground="white", font=("Segoe UI", 10))
            entry.pack(anchor="w", padx=10, pady=2)
            self.entries[field] = entry

    def generate_documents(self):
        output_dir = filedialog.askdirectory(title="Select Output Folder", initialdir=self.output_folder)
        if not output_dir:
            return

        self.output_folder = output_dir
        self.save_config()

        values = {key: entry.get() for key, entry in self.entries.items()}

        for filename in self.selected_templates:
            try:
                doc_path = os.path.join(self.template_folder, filename)
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

        def get_effective_font(run):
            if run.font.name:
                return run.font.name
            try:
                rfonts = run._element.xpath(".//w:rFonts")
                if rfonts:
                    return rfonts[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii")
            except:
                pass
            return None

        def split_runs_by_char(paragraph):
            chars = []
            for run in paragraph.runs:
                font_name = get_effective_font(run)
                font_size = run.font.size
                for c in run.text:
                    chars.append({
                        "char": c,
                        "bold": run.bold,
                        "italic": run.italic,
                        "underline": run.underline,
                        "font": font_name,
                        "size": font_size
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
                fmt = chars[start] if start < len(chars) else {
                    "bold": None, "italic": None, "underline": None,
                    "font": None, "size": None
                }
                new_chars.extend([{ 
                    "char": rc, 
                    "bold": fmt['bold'], 
                    "italic": fmt['italic'], 
                    "underline": fmt['underline'], 
                    "font": fmt['font'], 
                    "size": fmt['size'] 
                } for rc in replacement])
                last_idx = end

            new_chars.extend(chars[last_idx:])

            for run in paragraph.runs:
                run.text = ""

            new_run = None
            for ch in new_chars:
                if (
                    new_run is None or
                    new_run.bold != ch['bold'] or
                    new_run.italic != ch['italic'] or
                    new_run.underline != ch['underline'] or
                    (new_run.font.name != ch['font'] if ch['font'] else False) or
                    (new_run.font.size != ch['size'] if ch['size'] else False)
                ):
                    new_run = paragraph.add_run()
                    new_run.bold = ch['bold']
                    new_run.italic = ch['italic']
                    new_run.underline = ch['underline']
                    if ch['font']:
                        try:
                            new_run.font.name = ch['font']
                        except:
                            pass
                    if ch['size']:
                        try:
                            new_run.font.size = ch['size']
                        except:
                            pass
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
