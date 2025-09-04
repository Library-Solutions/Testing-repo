# ==================== Imports ====================
print("Started Auto code")
import os
import sys
import json
import socket
import logging
import threading
import requests
import base64
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import sleep
from tkinter import Tk, Canvas, Frame, Button, CENTER
from PIL import Image, ImageTk
from configparser import ConfigParser
from rich.console import Console
from rich.logging import RichHandler
from rich.traceback import install

import glob
import gdown
import pandas as pd
import shutil
import openpyxl

from googledriver import download_folder


# ==================== Setup ====================
sys.setrecursionlimit(5000)

SYSTEM_PLATFORM = not (sys.platform == "win32" or sys.platform == "darwin")

if SYSTEM_PLATFORM:
    import neopixel
    import board
    import digitalio

    os.system('vcgencmd display_power 0')
    pixels = neopixel.NeoPixel(board.D21, 60)
    paperSensor = digitalio.DigitalInOut(board.D5)
    paperSensor.direction = digitalio.Direction.INPUT


# ==================== Project Paths ====================
PROJECT_ROOT = os.path.dirname(os.path.realpath(__file__))
ICONS_ROOT = os.path.join(PROJECT_ROOT, 'icons')
LOG_FILE = os.path.join(PROJECT_ROOT, 'logData', 'system.log')
CONFIG_FILE = os.path.join(PROJECT_ROOT, 'systemConfig.ini')
EXCELLSHEET_ROOT = os.path.join(PROJECT_ROOT, "excellFiles")
DOCUMENTS_ROOT = os.path.join(PROJECT_ROOT, 'textDocuments')

SYSTEM_EXCEL_FILE_NAME = os.path.join(PROJECT_ROOT, "systemExcelFile.xlsx")
CIMMULATIVE_EXCEL_FILE_NAME = os.path.join(PROJECT_ROOT, "cumulativeExcelSheet.xlsx")

EXCEPTION_FOLDERS = {
    "excellFiles",
    "icons",
    "logData",
    "textDocuments"
}

DELETE_FOLDERS_FILES = [
    "excellFiles", "icons", "logData", "textDocuments",
    "autoUpdate.py", "autoUpdate.pyc", "pythonCode.py",
    "pythonCode.pyc", "systemConfig.ini", "iconList.ini", "pythonCode.py"
]


# ==================== Logging Setup ====================
console = Console()
install(show_locals=True)

def setup_logger(name, log_file):
    """Configure and return a logger instance with Rich and File handlers."""
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    formatter = logging.Formatter(
        "%(asctime)s :: %(levelname)-8s :: %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(formatter)

    console_handler = RichHandler(console=console)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

systemLogger = setup_logger("systemLogger", LOG_FILE)


# ==================== Input Data ====================
input_data = ""

with open("data.json", "r") as f:
    input_data = json.load(f)

data = input_data
print(data)
print(data[0].get("USAGE_DATA") )

# ==================== Config Parsing ====================
CONFIG_KEYS = [
    "SBN_ID", "CLIENT_NAME", "LOCATION", "CITY", "STATE", "COUNTRY",
    "RENTAL_OR_SUBSCRIPTION", "RENEWAL_DATE", "COBRANDING", "EVENT_TYPE",
    "EVENT_NUMBER", "QR_ENABLED", "QR_EMAILID", "QR_PAYMENT_URL", "SHEETID"
]

config_values = {}


def load_config():
    """Load and parse values from the INI config file."""
    config = ConfigParser()
    config.read(CONFIG_FILE)

    try:
        info = config["storyboxinfo"]
        for key in CONFIG_KEYS:
            config_values[key] = info.get(key)
        config_values["RENEWAL_DATE"] = int(config_values.get("RENEWAL_DATE", 0))
        config_values["EVENT_NUMBER"] = int(config_values.get("EVENT_NUMBER", 0))
    except KeyError as e:
        systemLogger.error(f"Missing section in config: {e}")

load_config()

SCRIPT_URL = f"https://script.google.com/macros/s/{config_values['SHEETID']}/exec"

# ==================== GUI ====================
WINDOW_WIDTH, WINDOW_HEIGHT = 800, 480

window = Tk()
window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
window.resizable(False, False)
window.overrideredirect(True)

update_page = Frame(window, width=WINDOW_WIDTH, height=WINDOW_HEIGHT, bg="white")
update_canvas = None
comment_text_id = None
image_id = None

def raise_frame(frame):
    """Bring the specified frame to the front and enable display."""
    os.system('vcgencmd display_power 1')
    frame.tkraise()
    frame.pack()


def quit_app():
    """Close the application and log the exit."""
    systemLogger.info("Exiting application.")
    window.destroy()
    #window.quit()
    sys.exit()


# ==================== System Actions ====================
def check_internet():
    """Check internet connectivity by pinging a reliable DNS."""
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=5)
        systemLogger.info("Internet connected.")
        return True
    except OSError as e:
        systemLogger.warning(f"No internet connection: {e}")
        return False


def launch_script(script_name):
    """Terminate current window and launch another Python script."""
    os.system('vcgencmd display_power 0')
    #window.destroy()
    #window.quit()
    os.system('sudo python3 '+os.path.join(PROJECT_ROOT, script_name))


# ==================== UI Actions ====================
def update_screen(image="Downloading_red.png", message="Getting information..."):
    """Display an update screen with an image and status message."""
    global update_canvas, comment_text_id, image_id
    try:
        update_canvas = Canvas(
            update_page, bg="white",
            width=WINDOW_WIDTH, height=WINDOW_HEIGHT,
            bd=0, highlightthickness=0
        )
        update_canvas.place(x=0, y=0)

        img = Image.open(os.path.join(ICONS_ROOT, image))
        update_img = ImageTk.PhotoImage(img)
        update_canvas.update_img = update_img  # Prevent garbage collection
        image_id = update_canvas.create_image(400, 220, image=update_img, anchor=CENTER)
        comment_text_id = update_canvas.create_text(
            400, 45, text=message, font="Candara 16 bold", fill="#652828"
        )

        quit_img = Image.open(os.path.join(ICONS_ROOT, "error_quit.png"))
        quit_btn_img = ImageTk.PhotoImage(quit_img)
        Button(
            update_canvas, image=quit_btn_img,
            bg="white", borderwidth=0, command=quit_app
        ).place(x=795, y=0)
        update_canvas.quit_button_img = quit_btn_img

        raise_frame(update_page)

    except Exception as e:
        systemLogger.error(f"UI initialization error: {e}")
        quit_app()


def update_message(text, image=None):
    """Update the message and optionally image on the update canvas."""
    def _update():
        try:
            if image:
                img = Image.open(os.path.join(ICONS_ROOT, image))
                img_tk = ImageTk.PhotoImage(img)
                update_canvas.update_img = img_tk
                update_canvas.itemconfig(image_id, image=img_tk)
            update_canvas.itemconfigure(comment_text_id, text=text)
        except Exception as e:
            systemLogger.error(f"Error updating message: {e}")

    window.after(0, _update)


# ==================== Cloud Communication ====================
def upload_file(file_type):
    """Upload a file to a remote server, encoded in base64 if needed."""
    try:
        now = datetime.now()
        prev_month = now - relativedelta(months=1)

        if file_type == "USAGE_DATA_EXCEL":
            fname = f"{config_values['SBN_ID']}_{prev_month.strftime('%m')}-{prev_month.year}.xls"
        elif file_type == "USAGE_DATA_EXCEL_BACKUP":
            fname = f"{config_values['SBN_ID']}_backup.log"
        elif file_type == "CONTENT_CROSSCHECK_FILE":
            fname = "cumulativeExcelSheet.xlsx"
        else:
            systemLogger.warning(f"Unknown file type: {file_type}")
            return

        if "USAGE" in file_type:
            fpath = os.path.join(EXCELLSHEET_ROOT, fname)
        else:
            fpath = os.path.join(PROJECT_ROOT, fname)

        if os.path.exists(fpath):
            if "USAGE" in file_type:
                update_message("Sending usage data...", "Downloading_UsageDataUpdateIcon.png")
            else:
                update_message("Sending content data...", "Downloading_UsageDataUpdateIcon.png")

            with open(fpath, "rb") as f:
                encoded = base64.urlsafe_b64encode(f.read())
        else:
            update_message("File Not Found...", "Downloading_UsageDataUpdateIcon.png")
            encoded = None
            fname = f"FileNotFound - {fname}"

        params = {
            "StoryboxId": config_values["SBN_ID"],
            "Configuration": file_type,
            "FileName": fname
        }

        response = requests.post(SCRIPT_URL, params=params, data=encoded if encoded else None)

        if response.status_code == 200:
            systemLogger.info(f"Uploaded {file_type} successfully.")
        else:
            systemLogger.error(f"Upload failed ({response.status_code}): {response.text}")

    except Exception as e:
        update_message(f"Upload error ({file_type})", "Downloading_RandonErrorIcon.png")
        systemLogger.error(f"Upload error ({file_type}): {e}")

# ==================== Content Update Logic ====================
def updateDownload_File(downloaderRun):
    """Threaded function that monitors live download progress."""
    try:
        while downloaderRun.is_set():
            try:
                folders = glob.glob(os.path.join(PROJECT_ROOT, '*/'))
                latest_folder = max(folders, key=os.path.getmtime)
                last_folder_name = os.path.basename(os.path.normpath(latest_folder))

                if last_folder_name not in EXCEPTION_FOLDERS:
                    file_count = len(os.listdir(latest_folder))
                    info_text = f"Downloading file number {file_count} from folder {last_folder_name}"
                    systemLogger.info(info_text)
                    update_message(info_text, "Downloading_ContentUpdateIcon.png")
                else:
                    systemLogger.info("Monitoring download folders - No update yet.")
            except Exception as e:
                systemLogger.warning(f"Folder monitoring error: {e}")
            sleep(1)
    except Exception as e:
        update_message("Critical error in updateDownload_File", "Downloading_RandonErrorIcon.png")
        systemLogger.error(f"Critical error in updateDownload_File: {e}")


def fileUpdates(CUMULATIVE_EXCEL_FILE_NAME):
    """Download and restore missing files as listed in the cumulative Excel."""
    missingFilePath = []

    try:
        cumulative_sheets = pd.read_excel(CUMULATIVE_EXCEL_FILE_NAME, sheet_name=None)
        sheet1 = cumulative_sheets.get("Sheet1")
        sheet2 = cumulative_sheets.get("Sheet2")

        if sheet1 is None or sheet2 is None:
            systemLogger.error("Missing required sheets in cumulative Excel.")
            return

        for idx, row in sheet1.iterrows():
            rel_path = row.get("path")
            file_name = row.get("file_name")
            version = row.get("version")
            full_path = os.path.join(PROJECT_ROOT, rel_path, file_name)

            if not os.path.exists(full_path):
                version_row = sheet2.loc[sheet2["v"] == version]
                if version_row.empty:
                    systemLogger.warning(f"No link found for version {version}")
                    continue

                version_url = version_row.iloc[0]["link"]
                download_dir = os.path.join(PROJECT_ROOT, version)

                downloaderRun = threading.Event()
                downloaderRun.set()

                thread = threading.Thread(target=updateDownload_File, args=(downloaderRun,))
                thread.start()

                download_folder(version_url, download_dir)

                downloaderRun.clear()
                thread.join()

                for idx2, row2 in sheet1.iterrows():
                    fname = row2["file_name"]
                    dest_path = row2["path"]
                    update_dir = os.path.join(PROJECT_ROOT, str(version))
                    src_file = os.path.join(update_dir, fname)

                    if os.path.exists(src_file):
                        target_dir = os.path.join(PROJECT_ROOT, dest_path)
                        os.makedirs(target_dir, exist_ok=True)
                        shutil.copy(src_file, target_dir)
                        missingFilePath.append(os.path.join(dest_path, fname))
                        systemLogger.info(f"Restored: {fname}")

    except Exception as e:
        update_message("Error during file update", "Downloading_RandonErrorIcon.png")
        systemLogger.error(f"Error during file update: {e}")

    try:
        if missingFilePath:
            with pd.ExcelWriter(CUMULATIVE_EXCEL_FILE_NAME, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
                pd.DataFrame(missingFilePath, columns=["Downloaded Files"]).to_excel(writer, sheet_name="DownloadedPath", index=False)
    except Exception as e:
        update_message("Failed to write DownloadedPath", "Downloading_RandonErrorIcon.png")
        systemLogger.warning(f"Failed to write DownloadedPath: {e}")

    upload_file("CONTENT_CROSSCHECK_FILE")


def fileCheck(CUMULATIVE_EXCEL_FILE_NAME, SYSTEM_EXCEL_FILE_NAME):
    """Compare the current file system with the cumulative record and remove obsolete files."""
    try:
        system_df = pd.read_excel(SYSTEM_EXCEL_FILE_NAME)
        cumulative_sheets = pd.read_excel(CUMULATIVE_EXCEL_FILE_NAME, sheet_name=None)

        sheet1 = cumulative_sheets.get("Sheet1")
        if sheet1 is None:
            systemLogger.error("Sheet1 not found in cumulative Excel.")
            return

        system_paths = system_df["Path"].tolist()
        cumulative_paths = [
            os.path.normpath(os.path.join(PROJECT_ROOT, row["path"], row["file_name"]))
            for _, row in sheet1.iterrows()
        ]

        deleted_paths = []
        for path in system_paths:
            norm_path = os.path.normpath(path)
            if norm_path not in cumulative_paths and os.path.exists(norm_path):
                os.remove(norm_path)
                deleted_paths.append(path)
                systemLogger.info(f"Deleted obsolete file: {path}")

        if deleted_paths:
            with pd.ExcelWriter(CUMULATIVE_EXCEL_FILE_NAME, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
                pd.DataFrame(deleted_paths, columns=["Deleted Files"]).to_excel(writer, sheet_name="DeletedPath", index=False)

    except Exception as e:
        update_message("Error during file check", "Downloading_RandonErrorIcon.png")
        systemLogger.error(f"Error during file check: {e}")

    fileUpdates(CUMULATIVE_EXCEL_FILE_NAME)


def contentUpdate(excelFileID):
    """Main entry point for content update. Downloads Excel and performs checks and sync."""
    try:
        update_message("Updating content", "Downloading_ContentUpdateIcon.png")
        excel_url = f"https://drive.google.com/uc?id={excelFileID}"

        gdown.download(
            url=excel_url,
            output=CIMMULATIVE_EXCEL_FILE_NAME,
            quiet=False,
            fuzzy=True,
            use_cookies=False
        )

        systemLogger.info("Generating list of story files from system directory.")
        directory = os.path.join(PROJECT_ROOT, "textDocuments")
        file_list = glob.glob(directory + "/**/*.jpg", recursive=True)
        file_names = [os.path.basename(f) for f in file_list]

        df = pd.DataFrame({"Path": file_list, "fileName": file_names})
        df.to_excel(SYSTEM_EXCEL_FILE_NAME, index=False)

        fileCheck(CIMMULATIVE_EXCEL_FILE_NAME, SYSTEM_EXCEL_FILE_NAME)

    except Exception as e:
        update_message("Content update failed", "Downloading_RandonErrorIcon.png")
        systemLogger.error(f"Content update failed: {e}")

# ==================== Update Logic ====================
def run_update():
    try:
        """Main function controlling update logic, content sync, and file uploads."""
        if not check_internet():
            update_message("Internet Not Connected", "Downloading_RandonErrorIcon.png")
            sleep(1)
            return launch_script("pythonCode.py")
        today = int(datetime.now().strftime("%Y%m%d"))
        renewal_due = int(data[0]["RENEWAL_DATE"]) <= today

        if not renewal_due:
            if data[0].get("USAGE_DATA", "YES") == "NO":
                upload_file("USAGE_DATA_EXCEL")
                upload_file("USAGE_DATA_EXCEL_BACKUP")

            if data[0].get("CONTENT_UPDATE", "YES") == "NO":
                contentUpdate(data[0]["CONTENT_UPDATE_EXCEL_LINK"])

            if SYSTEM_PLATFORM:
                launch_script("pythonCode.py")
            else:
                systemLogger.debug("Dev mode - Code running in Windows/macOS - skipping auto-launch after update.")
                quit_app()
        else:
            if SYSTEM_PLATFORM:
                launch_script("pythonCode.py")
            else:
                systemLogger.debug("Dev mode - Code running in Windows/macOS - renewal overdue, skipping launch.")
                quit_app()
    except Exception as e:
        print(f"Exception in run_update2: {e}", exc_info=True)
        systemLogger.error(f"Exception in run_update2: {e}", exc_info=True)


# ==================== Start ====================
if __name__ == "__main__":
	
    update_screen()
    thread = threading.Thread(target=run_update, daemon=True)
    thread.start()
    window.mainloop()
