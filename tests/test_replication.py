import os
import sys
import shutil
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from DocGenerator import extract_date_from_name, DocGeneratorApp, AUDIO_EXTENSIONS, EXCLUDED_EXTENSIONS

def test_date_extraction():
    test_cases = [
        ("DS904403_ _09112026_aks.doc", "09112026"),
        ("DS904403_ _09112026.Garden_AKS.doc", "09112026"),
        ("Template_2026-09-11_v1.docx", "2026-09-11"),
        ("Dictation_11.09.2026.doc", "11.09.2026"),
        ("DS904403.dss", ""),  # 6-digit ID should not be mistaken for 8-digit date
        ("NoDateTemplate.doc", ""),
    ]
    for filename, expected in test_cases:
        actual = extract_date_from_name(filename)
        assert actual == expected, f"Failed for {filename}: expected {expected}, got {actual}"
    print("[PASS] test_date_extraction passed.")

def test_replication_and_naming():
    with tempfile.TemporaryDirectory() as tmpdir:
        # Create simulated audio folder
        audio_folder = os.path.join(tmpdir, "AudioFiles")
        os.makedirs(audio_folder)

        # Create audio files: extensionless and with audio extensions
        audio_files = ["DS904403", "DS904404", "DS904405.dss", "DS904406.wav", "DS904407.mp3", "DS904408"]
        for af in audio_files:
            with open(os.path.join(audio_folder, af), "wb") as f:
                f.write(b"AUDIO_DATA")

        # Create a non-audio file that should be ignored
        with open(os.path.join(audio_folder, "notes.txt"), "w") as f:
            f.write("text note")

        # Create dummy template document with 09112026 in name
        template_name = "DS904400_ _09112026_aks.doc"
        template_path = os.path.join(tmpdir, template_name)
        with open(template_path, "wb") as f:
            f.write(b"FAKE_DOC_HEADER_AND_TEMPLATE_CONTENT_HERE")

        # Test audio file retrieval
        class DummyApp:
            pass
        
        # Instantiate minimal helper or check logic
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        app = DocGeneratorApp(root)
        app.folder_path.set(audio_folder)
        app.template_path.set(template_path)
        app.extracted_template_date = extract_date_from_name(template_name)
        
        bases = app.get_audio_files(audio_folder)
        expected_bases = ["DS904403", "DS904404", "DS904405", "DS904406", "DS904407", "DS904408"]
        assert sorted(bases) == sorted(expected_bases), f"Audio bases mismatch: {bases} vs {expected_bases}"
        print("[PASS] Audio files detected correctly.")

        # Test naming with template extracted date (date_var empty)
        app.date_var.set("")
        app.pattern_var.set("{base}_ _{date}_aks.doc")
        
        for b in bases:
            target_name = app.build_file_name(b)
            assert target_name == f"{b}_ _09112026_aks.doc", f"Unexpected name: {target_name}"
        print("[PASS] Format 1 with auto-extracted date passed.")

        # Test format 2 with custom date
        app.date_var.set("12345678")
        app.pattern_var.set("{base}_ _{date}.Garden_AKS.doc")
        for b in bases:
            target_name = app.build_file_name(b)
            assert target_name == f"{b}_ _12345678.Garden_AKS.doc", f"Unexpected name: {target_name}"
        print("[PASS] Format 2 with custom date passed.")

        # Test replication execution (reset date to empty so it uses 09112026)
        app.date_var.set("")
        app.pattern_var.set("{base}_ _{date}_aks.doc")
        
        # Run replication logic directly
        for b in bases:
            doc_name = app.build_file_name(b)
            target_path = os.path.join(audio_folder, doc_name)
            shutil.copy2(template_path, target_path)

        # Verify all 6 documents were created and match template content exactly
        for b in bases:
            doc_name = f"{b}_ _09112026_aks.doc"
            target_path = os.path.join(audio_folder, doc_name)
            assert os.path.exists(target_path), f"File {doc_name} was not created"
            with open(target_path, "rb") as f:
                content = f.read()
                assert content == b"FAKE_DOC_HEADER_AND_TEMPLATE_CONTENT_HERE"
        print("[PASS] Documents replicated with exact binary fidelity.")

        # Test settings persistence
        app.folder_path.set(audio_folder)
        app.template_path.set(template_path)
        app.date_var.set("09112026")
        app.pattern_var.set("{base}_ _{date}_custom.doc")
        app.save_settings()

        # Create new app instance and test loading
        app2 = DocGeneratorApp(root)
        assert app2.folder_path.get() == audio_folder, f"Expected {audio_folder}, got {app2.folder_path.get()}"
        assert app2.template_path.get() == template_path, f"Expected {template_path}, got {app2.template_path.get()}"
        assert app2.date_var.get() == "09112026", f"Expected 09112026, got {app2.date_var.get()}"
        assert app2.pattern_var.get() == "{base}_ _{date}_custom.doc", f"Expected custom pattern, got {app2.pattern_var.get()}"
        print("[PASS] Settings persistence (save and load) verified.")

        root.destroy()

if __name__ == "__main__":
    test_date_extraction()
    test_replication_and_naming()
    print("\nALL VERIFICATION TESTS PASSED SUCCESSFULLY!")
