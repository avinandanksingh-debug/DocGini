import os
import re
import sys
import json
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

# Enable Windows High-DPI awareness for crisp rendering
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# Audio file extensions (including Olympus/Philips dictations and raw extensionless files)
AUDIO_EXTENSIONS = {
    ".dss",
    ".ds2",
    ".wav",
    ".mp3",
    ".m4a",
    ".wma",
    ".aac",
    ".flac",
    ".ogg",
    ".aiff",
    ".amr",
    "",  # Handles files with no extension (e.g. Olympus/Philips raw IDs)
}

# Explicitly excluded extensions that should never be treated as audio files
EXCLUDED_EXTENSIONS = {
    ".doc",
    ".docx",
    ".dot",
    ".dotx",
    ".txt",
    ".pdf",
    ".rtf",
    ".py",
    ".pyw",
    ".bat",
    ".cmd",
    ".ps1",
    ".exe",
    ".dll",
    ".zip",
    ".rar",
    ".7z",
    ".tar",
    ".gz",
    ".ini",
    ".lnk",
    ".ico",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".json",
    ".xml",
    ".csv",
    ".xlsx",
    ".xls",
}


def extract_date_from_name(filename: str) -> str:
    """
    Intelligently extracts a date string from a file or template name.
    Supports:
      - 8 consecutive digits (e.g. 09112026, 20260911, 11092026)
      - Delimited dates (e.g. 09-11-2026, 2026-09-11, 09.11.2026)
    """
    base_name = os.path.splitext(filename)[0]
    # 1. Look for 8 consecutive digits not adjacent to other digits
    m = re.search(r"(?<!\d)(\d{8})(?!\d)", base_name)
    if m:
        return m.group(1)
    # 2. Look for hyphen/dot/underscore separated date formats
    m = re.search(r"(?<!\d)(\d{2,4}[-._]\d{2}[-._]\d{2,4})(?!\d)", base_name)
    if m:
        return m.group(1)
    return ""


def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def get_config_file_path() -> str:
    """Returns path to config.json. Supports portable mode and %APPDATA%."""
    # Check if a config.json is present in the app's directory (portable mode)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    local_config = os.path.join(exe_dir, "config.json")
    if os.path.exists(local_config):
        return local_config

    # Default location in %APPDATA%\DocGini\config.json
    appdata = os.getenv("APPDATA")
    if appdata:
        config_dir = os.path.join(appdata, "DocGini")
        try:
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, "config.json")
        except Exception:
            pass

    return local_config


class DocGeneratorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DocGini by Aks Labs — Batch Word Document Generator")
        self.root.geometry("780x750")
        self.root.minsize(680, 600)

        self.extracted_template_date = ""

        # Set application icon
        self.setup_icon()

        # Configure TTK style
        self.setup_styles()

        # Build UI layout
        self.build_ui()

        # Load saved settings & paths from previous session
        self.load_settings()

        # Auto-save settings on close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_icon(self):
        """Load 3D rounded app icon if available."""
        ico_path = get_resource_path(os.path.join("assets", "icon.ico"))
        png_path = get_resource_path(os.path.join("assets", "icon.png"))

        try:
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
            if os.path.exists(png_path):
                self.icon_photo = tk.PhotoImage(file=png_path)
                self.root.iconphoto(True, self.icon_photo)
        except Exception:
            pass

    def setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Custom colors and styles
        style.configure("TLabel", font=("Segoe UI", 9))
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1e3a8a")
        style.configure("Subheader.TLabel", font=("Segoe UI", 8), foreground="#4b5563")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground="#1f2937")
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), padding=6)
        style.configure("TButton", font=("Segoe UI", 9), padding=5)

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # App Header Banner
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        title_lbl = ttk.Label(header_frame, text="DocGini by Aks Labs", style="Header.TLabel")
        title_lbl.pack(anchor=tk.W)
        sub_lbl = ttk.Label(
            header_frame,
            text="Replicate template documents for all audio files with custom naming & date presets.",
            style="Subheader.TLabel"
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # --- STEP 1: FOLDER SELECTION ---
        folder_frame = ttk.LabelFrame(main_frame, text=" 1. Select Folder with Audio Files ", padding="10", style="Section.TLabelframe")
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        self.folder_path = tk.StringVar()
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path, font=("Segoe UI", 9))
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(folder_frame, text="Browse Folder...", command=self.browse_folder).pack(side=tk.RIGHT)

        # --- STEP 2: TEMPLATE SELECTION ---
        template_frame = ttk.LabelFrame(main_frame, text=" 2. Select Template Document (*.doc, *.docx) ", padding="10", style="Section.TLabelframe")
        template_frame.pack(fill=tk.X, pady=(0, 10))

        self.template_path = tk.StringVar()
        template_entry = ttk.Entry(template_frame, textvariable=self.template_path, font=("Segoe UI", 9))
        template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(template_frame, text="Browse Template...", command=self.browse_template).pack(side=tk.RIGHT)

        self.template_info_var = tk.StringVar(value="No template selected. New documents will be exact clones of this template.")
        self.template_info_lbl = ttk.Label(
            template_frame,
            textvariable=self.template_info_var,
            font=("Segoe UI", 8, "italic"),
            foreground="#6b7280"
        )
        self.template_info_lbl.pack(anchor=tk.W, pady=(5, 0))

        # --- STEP 3: CONFIGURATION (Naming Pattern & Date) ---
        config_frame = ttk.LabelFrame(main_frame, text=" 3. Naming Pattern & Date Settings ", padding="10", style="Section.TLabelframe")
        config_frame.pack(fill=tk.X, pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)

        # Date input
        ttk.Label(config_frame, text="Date String:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=4)
        date_subframe = ttk.Frame(config_frame)
        date_subframe.grid(row=0, column=1, sticky=tk.EW, pady=4, padx=(8, 0))

        self.date_var = tk.StringVar()
        self.date_entry = ttk.Entry(date_subframe, textvariable=self.date_var, width=18, font=("Segoe UI", 9))
        self.date_entry.pack(side=tk.LEFT)
        self.date_var.trace_add("write", lambda *args: self.preview_names(silent=True))

        self.date_hint_var = tk.StringVar(value="(Optional: leave blank to keep date from template name)")
        date_hint = ttk.Label(
            date_subframe,
            textvariable=self.date_hint_var,
            font=("Segoe UI", 8, "italic"),
            foreground="#6b7280"
        )
        date_hint.pack(side=tk.LEFT, padx=(10, 0))

        # Quick date helper button (Today)
        ttk.Button(
            date_subframe,
            text="Use Today",
            command=lambda: self.date_var.set(datetime.now().strftime("%m%d%Y")),
            width=10
        ).pack(side=tk.RIGHT)

        # Preset format dropdown
        ttk.Label(config_frame, text="Format Presets:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=4)
        self.presets = [
            "{base}_ _{date}_aks.doc",
            "{base}_ _{date}.Garden_AKS.doc",
            "{base}_ _{date}_aks.docx",
            "{base}_ _{date}.Garden_AKS.docx",
            "Custom Pattern...",
        ]
        self.preset_var = tk.StringVar(value=self.presets[0])
        self.preset_cb = ttk.Combobox(
            config_frame,
            textvariable=self.preset_var,
            values=self.presets,
            state="readonly",
            font=("Segoe UI", 9),
        )
        self.preset_cb.grid(row=1, column=1, sticky=tk.EW, pady=4, padx=(8, 0))
        self.preset_cb.bind("<<ComboboxSelected>>", self.on_preset_change)

        # Active Template pattern box
        ttk.Label(config_frame, text="Active Pattern:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky=tk.W, pady=4)
        self.pattern_var = tk.StringVar(value=self.presets[0])
        self.pattern_entry = ttk.Entry(
            config_frame,
            textvariable=self.pattern_var,
            font=("Segoe UI", 9),
        )
        self.pattern_entry.grid(row=2, column=1, sticky=tk.EW, pady=4, padx=(8, 0))
        self.pattern_var.trace_add("write", lambda *args: self.preview_names(silent=True))

        hint = ttk.Label(
            config_frame,
            text="Tags: {base} = Audio File Name, {date} = Date string (or extracted from template).",
            font=("Segoe UI", 8, "italic"),
            foreground="#4b5563",
        )
        hint.grid(row=3, column=1, sticky=tk.W, padx=(8, 0), pady=(2, 0))

        # --- STEP 4: ACTIONS & PROGRESS ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(2, 8))

        ttk.Button(btn_frame, text="🔍 Preview Files", command=lambda: self.preview_names(silent=False)).pack(side=tk.LEFT, padx=(0, 6))
        self.gen_btn = ttk.Button(
            btn_frame,
            text="⚡ Generate Documents",
            command=self.generate_documents,
            style="Accent.TButton",
        )
        self.gen_btn.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(btn_frame, text="📂 Open Folder", command=self.open_target_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text="🧹 Clear Log", command=self.clear_log).pack(side=tk.RIGHT)

        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        # --- STEP 5: PREVIEW / LOG WINDOW ---
        log_frame = ttk.LabelFrame(main_frame, text=" Preview / Output Log ", padding="8", style="Section.TLabelframe")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#fdfdfd",
            fg="#1f2937",
            padx=5,
            pady=5,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Configure color tags for log
        self.log_text.tag_config("title", font=("Consolas", 9, "bold"), foreground="#1e3a8a")
        self.log_text.tag_config("created", foreground="#047857", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("skipped", foreground="#b45309")
        self.log_text.tag_config("error", foreground="#b91c1c", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("info", foreground="#374151")

        self.log_text.insert(
            tk.END,
            "Welcome to DocGini!\n"
            "1. Select the folder containing your audio files.\n"
            "2. Select your Word template document (*.doc / *.docx).\n"
            "3. Enter a date or leave blank to keep the date from the template name.\n"
            "4. Click 'Preview Files' to verify, then 'Generate Documents'.\n\n",
            "info"
        )

    def browse_folder(self):
        selected = filedialog.askdirectory(title="Select Folder Containing Audio Files")
        if selected:
            self.folder_path.set(selected)
            self.preview_names(silent=True)
            self.save_settings()

    def apply_template(self, selected: str, auto_adjust_pattern: bool = True):
        self.template_path.set(selected)
        filename = os.path.basename(selected)
        ext = os.path.splitext(filename)[1].lower()

        # Attempt to extract date from template name
        extracted = extract_date_from_name(filename)
        self.extracted_template_date = extracted

        if extracted:
            self.template_info_var.set(f"Selected: {filename}  |  Extracted Date: {extracted}")
            self.date_hint_var.set(f"(Blank = will use template date: {extracted})")
        else:
            self.template_info_var.set(f"Selected: {filename}  |  (No date pattern detected in template name)")
            self.date_hint_var.set("(Optional: specify custom date)")

        # If auto_adjust_pattern is True, auto-match the template extension in active pattern
        if auto_adjust_pattern and ext in [".doc", ".docx"]:
            current_pat = self.pattern_var.get().strip()
            if ext == ".doc" and current_pat.endswith(".docx"):
                self.pattern_var.set(current_pat[:-1])
            elif ext == ".docx" and current_pat.endswith(".doc"):
                self.pattern_var.set(current_pat + "x")

    def browse_template(self):
        selected = filedialog.askopenfilename(
            title="Select Template Document",
            filetypes=[
                ("Word Documents (*.doc;*.docx)", "*.doc;*.docx"),
                ("Legacy Word Document (*.doc)", "*.doc"),
                ("Word Document (*.docx)", "*.docx"),
                ("All Files (*.*)", "*.*"),
            ]
        )
        if selected:
            self.apply_template(selected, auto_adjust_pattern=True)
            self.preview_names(silent=True)
            self.save_settings()

    def on_preset_change(self, event=None):
        val = self.preset_var.get()
        if val != "Custom Pattern...":
            # If a template is loaded, adjust extension to match template if applicable
            t_path = self.template_path.get().strip()
            if t_path:
                t_ext = os.path.splitext(t_path)[1].lower()
                if t_ext == ".doc" and val.endswith(".docx"):
                    val = val[:-1]
                elif t_ext == ".docx" and val.endswith(".doc"):
                    val = val + "x"
            self.pattern_var.set(val)
        self.preview_names(silent=True)
        self.save_settings()

    def load_settings(self):
        """Loads and restores remembered settings and file paths from config.json."""
        config_file = get_config_file_path()
        if not os.path.exists(config_file):
            return

        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Restore template first (if file exists)
            saved_template = data.get("template_path", "")
            if saved_template and os.path.isfile(saved_template):
                self.apply_template(saved_template, auto_adjust_pattern=False)

            # Restore pattern & preset
            saved_preset = data.get("preset", "")
            if saved_preset in self.presets:
                self.preset_var.set(saved_preset)

            saved_pattern = data.get("pattern", "")
            if saved_pattern:
                self.pattern_var.set(saved_pattern)

            # Restore date string
            saved_date = data.get("date_string", "")
            if saved_date:
                self.date_var.set(saved_date)

            # Restore folder path (if folder exists)
            saved_folder = data.get("folder_path", "")
            if saved_folder and os.path.isdir(saved_folder):
                self.folder_path.set(saved_folder)
                self.preview_names(silent=True)

            # Restore window size/position
            saved_geom = data.get("geometry", "")
            if saved_geom:
                self.root.geometry(saved_geom)

        except Exception:
            pass

    def save_settings(self):
        """Saves current paths, pattern, date, and geometry to config.json."""
        config_file = get_config_file_path()
        try:
            data = {
                "folder_path": self.folder_path.get().strip(),
                "template_path": self.template_path.get().strip(),
                "date_string": self.date_var.get().strip(),
                "preset": self.preset_var.get(),
                "pattern": self.pattern_var.get().strip(),
                "geometry": self.root.geometry(),
            }
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def on_close(self):
        """Handles window closing event."""
        self.save_settings()
        self.root.destroy()

    def get_audio_files(self, folder: str):
        if not folder or not os.path.isdir(folder):
            return []
        
        try:
            entries = os.listdir(folder)
        except Exception:
            return []

        base_names = []
        for entry in entries:
            full_path = os.path.join(folder, entry)
            # Must be a file, not a directory
            if not os.path.isfile(full_path):
                continue
            # Ignore hidden or temporary files (e.g. ~$Doc.docx, .DS_Store, Thumbs.db)
            if entry.startswith("~$") or entry.startswith("."):
                continue

            name, ext = os.path.splitext(entry)
            ext_lower = ext.lower()

            # Exclude known non-audio formats
            if ext_lower in EXCLUDED_EXTENSIONS:
                continue

            # Must match audio extension or be an extensionless ID
            if ext_lower in AUDIO_EXTENSIONS:
                if name and name not in base_names:
                    base_names.append(name)

        base_names.sort()
        return base_names

    def get_effective_date(self) -> str:
        """Determines date: user input takes precedence, then template extracted date."""
        custom_date = self.date_var.get().strip()
        if custom_date:
            return custom_date
        return self.extracted_template_date

    def build_file_name(self, base_name: str) -> str:
        template = self.pattern_var.get().strip()
        date_str = self.get_effective_date()
        return template.replace("{base}", base_name).replace("{date}", date_str)

    def preview_names(self, silent=False):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            if not silent:
                messagebox.showwarning("Missing Folder", "Please select a valid folder containing audio files.")
            return

        base_names = self.get_audio_files(folder)
        self.log_text.delete("1.0", tk.END)

        if not base_names:
            self.log_text.insert(tk.END, f"Folder: {folder}\nNo audio files found matching criteria.\n", "info")
            return

        effective_date = self.get_effective_date()
        date_source = f"'{effective_date}'" if effective_date else "[NONE]"
        if not self.date_var.get().strip() and self.extracted_template_date:
            date_source += " (Extracted from template)"
        elif self.date_var.get().strip():
            date_source += " (User specified)"

        self.log_text.insert(tk.END, f"Folder: {folder}\n", "title")
        self.log_text.insert(tk.END, f"Audio Files Found: {len(base_names)} | Date in Use: {date_source}\n", "info")
        self.log_text.insert(tk.END, "-" * 75 + "\n", "info")

        template_selected = bool(self.template_path.get().strip() and os.path.isfile(self.template_path.get().strip()))

        for b in base_names:
            doc_name = self.build_file_name(b)
            target_path = os.path.join(folder, doc_name)
            if os.path.exists(target_path):
                status = "[EXISTS - Will Skip]"
                tag = "skipped"
            else:
                status = "[READY TO REPLICATE]" if template_selected else "[READY]"
                tag = "info"
            self.log_text.insert(tk.END, f"{b:<20} -> {doc_name:<40} {status}\n", tag)

    def generate_documents(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Missing Folder", "Please select a valid folder first.")
            return

        template_file = self.template_path.get().strip()
        if not template_file or not os.path.isfile(template_file):
            messagebox.showwarning(
                "Missing Template",
                "Please select a valid template Word document (*.doc / *.docx) to replicate."
            )
            return

        base_names = self.get_audio_files(folder)
        if not base_names:
            messagebox.showwarning("No Files", "No matching audio files to process in the selected folder.")
            return

        # Confirm if date is missing
        effective_date = self.get_effective_date()
        if not effective_date:
            proceed = messagebox.askyesno(
                "Date is Empty",
                "No date was entered and no date could be extracted from the template name.\n\n"
                "Do you want to proceed anyway with an empty date field?"
            )
            if not proceed:
                self.date_entry.focus_set()
                return

        self.log_text.delete("1.0", tk.END)
        self.log_text.insert(tk.END, "Starting batch document replication...\n", "title")
        self.log_text.insert(tk.END, f"Template: {os.path.basename(template_file)}\n", "info")
        self.log_text.insert(tk.END, f"Date: {effective_date or '[BLANK]'}\n", "info")
        self.log_text.insert(tk.END, "=" * 75 + "\n\n", "info")

        created_count = 0
        skipped_count = 0
        error_count = 0
        total = len(base_names)

        self.progress_var.set(0)
        self.root.update_idletasks()

        for idx, b in enumerate(base_names, start=1):
            doc_name = self.build_file_name(b)
            target_path = os.path.join(folder, doc_name)

            if os.path.exists(target_path):
                self.log_text.insert(tk.END, f"[SKIPPED (Already exists)]: {doc_name}\n", "skipped")
                skipped_count += 1
            else:
                try:
                    # Replicate template with 100% exact fidelity (styles, headers, formatting, binary structure)
                    shutil.copy2(template_file, target_path)
                    self.log_text.insert(tk.END, f"[REPLICATED & CREATED]: {doc_name}\n", "created")
                    created_count += 1
                except Exception as e:
                    self.log_text.insert(tk.END, f"[ERROR creating {doc_name}]: {e}\n", "error")
                    error_count += 1

            # Update progress
            self.progress_var.set((idx / total) * 100)
            self.root.update_idletasks()
            self.log_text.see(tk.END)

        self.log_text.insert(tk.END, "\n" + "=" * 75 + "\n", "info")
        summary_msg = f"Done! Created: {created_count} file(s) | Skipped: {skipped_count} file(s) | Errors: {error_count}"
        self.log_text.insert(tk.END, summary_msg + "\n", "title")
        self.log_text.see(tk.END)

        self.save_settings()

        if error_count == 0:
            messagebox.showinfo("Success", f"Document generation complete!\n\nCreated: {created_count}\nSkipped: {skipped_count}")
        else:
            messagebox.showwarning("Completed with Errors", f"Completed with some errors.\n\nCreated: {created_count}\nSkipped: {skipped_count}\nErrors: {error_count}")

    def open_target_folder(self):
        folder = self.folder_path.get().strip()
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Invalid Folder", "Please select a valid folder first.")

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)


def main():
    root = tk.Tk()
    app = DocGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
