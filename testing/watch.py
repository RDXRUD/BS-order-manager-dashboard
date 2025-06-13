import time
import subprocess
import os
import platform
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
from datetime import datetime

# ==== CONFIG ====
BASE_DIR = Path(__file__).parent.resolve()
WATCH_FOLDER = BASE_DIR / "data files" / "Purchase Order"
PDF_OUTPUT_FOLDER = BASE_DIR / "data files" / "Pdf Files"
SCRIPT_PATH = BASE_DIR / "script.py"
LOG_FILE = BASE_DIR / "conversion_log.txt"

PDF_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# ==== LOGGING SETUP ====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, mode="w",encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# Track both mod time and last handled time
file_mod_times = {}  # {path: (mod_time, last_handled_time)}

MIN_SECONDS_BETWEEN_RUNS = 5  # Cooldown window to avoid duplicate triggers

def open_pdf(pdf_path):
    try:
        if platform.system() == "Darwin":  # macOS
            subprocess.run(["open", str(pdf_path)])
        elif platform.system() == "Windows":
            os.startfile(str(pdf_path))
        else:  # Linux
            subprocess.run(["xdg-open", str(pdf_path)])
        logging.info(f"Opened PDF: {pdf_path}")
    except Exception as e:
        logging.warning(f"Could not open PDF: {e}")

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
            now = time.time()
        except FileNotFoundError:
            return
            
        if file_path not in file_mod_times or current_mod_time != file_mod_times[file_path]:
            print(file_mod_times)
            file_mod_times[file_path] = current_mod_time  # Update last seen mod time
            
            logging.info(f"📄 Valid update: {file_path.name}")

            try:
                today_str = datetime.now().strftime("%d-%m-%Y")
                dated_output_folder = PDF_OUTPUT_FOLDER / today_str
                dated_output_folder.mkdir(parents=True, exist_ok=True)

                pdf_filename = dated_output_folder / f"{file_path.stem}_{today_str}.pdf"

                subprocess.run(["python", str(SCRIPT_PATH), str(file_path), str(pdf_filename)], check=True)

                logging.info(f"✅ PDF created: {pdf_filename}")
                open_pdf(pdf_filename)

            except subprocess.CalledProcessError as e:
                logging.error(f"❌ Script failed for {file_path.name}: {e}")
            except Exception as e:
                logging.exception(f"💥 Unexpected error for {file_path.name}: {e}")

if __name__ == "__main__":
    logging.info(f"📁 Watching: {WATCH_FOLDER}")
    for file_path in WATCH_FOLDER.glob("*.xls*"):
        if not file_path.name.startswith("~$"):
            try:
                mod_time = file_path.stat().st_mtime
                file_mod_times[file_path] = (mod_time, 0)  # mod_time + dummy last_handled_time
                logging.info(f"🗂️ Preloaded: {file_path.name} with mod_time {mod_time}")
            except Exception as e:
                logging.warning(f"⚠️ Failed to preload {file_path.name}: {e}")
    event_handler = ExcelChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, str(WATCH_FOLDER), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logging.info("👋 Stopped watching.")
    observer.join()
