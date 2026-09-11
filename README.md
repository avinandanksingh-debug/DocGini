# DocGini by Aks Labs — Batch Word Document Generator

A desktop application designed to automate the creation and naming of multiple MS Word documents (`.doc` and `.docx`) corresponding to audio files in a folder, replicated directly from a preselected Word template document.

---

## 📦 Releases & Downloads

All release binaries are located in the [`release/`](file:///d:/Projects/DocGini/release/) directory:

1. **Standalone Windows Installer**:
   - [`DocGini_Setup_v1.0.exe`](file:///d:/Projects/DocGini/release/DocGini_Setup_v1.0.exe) (~14 MB)
   - Full setup wizard created with Inno Setup.
   - Installs Start Menu and Desktop shortcuts with the 3D rounded icon, supports clean uninstallation.
   - Requires no Python installation.

2. **Portable Edition**:
   - Standalone Executable: [`release/DocGini_Portable/DocGini.exe`](file:///d:/Projects/DocGini/release/DocGini_Portable/DocGini.exe) (~12 MB)
   - ZIP Archive: [`DocGini_v1.0_Portable.zip`](file:///d:/Projects/DocGini/release/DocGini_v1.0_Portable.zip) (~13 MB)
   - Runs instantly on any Windows PC without installation or admin privileges.
   - Built-in embedded 3D icon.

---

## ✨ Key Features

1. **Exact Template Replication**:
   - Every document is cloned directly from your preselected template document (`*.doc` or `*.docx`).
   - Retains 100% of the original formatting, headers, footers, tables, fonts, and styles byte-for-byte.

2. **Intelligent Date Detection**:
   - Automatically detects date sections in the template file name (e.g., `09112026` from `DS904400_ _09112026_aks.doc`).
   - If you leave the **Date String** field blank, the app preserves and uses the date extracted from your template document.
   - You can also type any custom date string (e.g., `09112026`) or click **"Use Today"** for today's date.

3. **Audio File Identification**:
   - Automatically identifies Olympus/Philips dictation files and audio formats:
     `.dss`, `.ds2`, `.wav`, `.mp3`, `.m4a`, `.wma`, `.aac`, `.flac`, `.ogg`, `.amr`, `.aiff`
   - Handles extensionless audio files (e.g. Olympus raw dictation files named `DS904403`, `DS904404`).
   - Ignores existing documents (`.doc`, `.docx`, `.pdf`, `.txt`), system files (`~$*`, `desktop.ini`), and scripts.

4. **Built-in Presets & Custom Formatting**:
   - `{base}_ _{date}_aks.doc`
   - `{base}_ _{date}.Garden_AKS.doc`
   - `{base}_ _{date}_aks.docx`
   - `{base}_ _{date}.Garden_AKS.docx`
   - Custom pattern support using `{base}` (Audio ID) and `{date}` (Date string).

5. **Live Preview & Collision Protection**:
   - Preview all target filenames before creating anything.
   - Existing documents are safely skipped to prevent accidental overwriting.

6. **3D Rounded App Icon & Clean Desktop Experience**:
   - Custom 3D squircle-style icon for the application window and Windows taskbar.
   - Embedded into the standalone executable and installer.

---

## 🚀 How to Run

### Method 1: Run Installer (Recommended)
Double-click [`release/DocGini_Setup_v1.0.exe`](file:///d:/Projects/DocGini/release/DocGini_Setup_v1.0.exe) to install DocGini with Start Menu & Desktop shortcuts.

### Method 2: Portable Executable
Double-click [`release/DocGini_Portable/DocGini.exe`](file:///d:/Projects/DocGini/release/DocGini_Portable/DocGini.exe) or unzip [`release/DocGini_v1.0_Portable.zip`](file:///d:/Projects/DocGini/release/DocGini_v1.0_Portable.zip) to run without installing.

### Method 3: Python Source
Double-click [`run.bat`](file:///d:/Projects/DocGini/run.bat) or run:
```powershell
python DocGenerator.pyw
```

---

## 📋 Step-by-Step Workflow

1. **Select Folder**:
   - Click **"Browse Folder..."** and pick the directory containing your audio files (`DS904403`, `DS904404`, etc.).
2. **Select Template Document**:
   - Click **"Browse Template..."** and choose your `.doc` or `.docx` template.
   - The app will automatically read any date in the template's filename and display it.
3. **Set Date & Pattern**:
   - Leave the date field blank to use the date from the template, or type a custom date (e.g. `09112026`).
   - Select one of your presets from the dropdown, or type a custom pattern.
4. **Preview & Generate**:
   - Click **"🔍 Preview Files"** to inspect the source-to-target document mapping.
   - Click **"⚡ Generate Documents"** to create all files in seconds.
   - Click **"📂 Open Folder"** to view your new documents in Windows File Explorer.
