import time
import subprocess
import os
import platform
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
from datetime import datetime

# ==== CONFIG ====
BASE_DIR = Path(__file__).parent.resolve()
WATCH_FOLDER = BASE_DIR / "data files" / "Purchase Order"
PDF_OUTPUT_FOLDER = BASE_DIR / "data files" / "Pdf Files"
SCRIPT_PATH = BASE_DIR / "script.py"

PDF_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# Keep track of file modification times
file_mod_times = {}

def open_pdf(pdf_path):
    try:
        if platform.system() == "Darwin":  # macOS
            subprocess.run(["open", str(pdf_path)])
        elif platform.system() == "Windows":
            os.startfile(str(pdf_path))
        else:  # Linux
            subprocess.run(["xdg-open", str(pdf_path)])
    except Exception as e:
        print(f"⚠️ Could not open PDF: {e}")

class ExcelChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        self._handle_event(event)

    def on_created(self, event):
        self._handle_event(event)

    def _handle_event(self, event):
        if event.is_directory:
            return

        file_path = Path(event.src_path)
        if file_path.suffix.lower() not in [".xlsx", ".xls"]:
            return
        if file_path.name.startswith("~$"):  # Ignore temp Excel files
            return

        try:
            current_mod_time = file_path.stat().st_mtime
        except FileNotFoundError:
            return

        if file_path not in file_mod_times or current_mod_time != file_mod_times[file_path]:
            print(f"[WATCHDOG] Detected real update: {file_path.name}")
            file_mod_times[file_path] = current_mod_time

            try:
                # 🗓️ Get today's date
                today_str = datetime.now().strftime("%d-%m-%Y")
                dated_output_folder = PDF_OUTPUT_FOLDER / today_str
                dated_output_folder.mkdir(parents=True, exist_ok=True)

                # 📝 PDF file name with date
                pdf_filename = dated_output_folder / f"{file_path.stem}_{today_str}.pdf"

                # 🛠️ Run script with full PDF path
                subprocess.run(["python", str(SCRIPT_PATH), str(file_path), str(pdf_filename)], check=True)

                print("✅ PDF created successfully.")
                open_pdf(pdf_filename)

            except subprocess.CalledProcessError as e:
                print(f"[ERROR] Script failed: {e}")

if __name__ == "__main__":
    print(f"📁 Watching: {WATCH_FOLDER}")
    event_handler = ExcelChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, str(WATCH_FOLDER), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n👋 Stopped watching.")
    observer.join()