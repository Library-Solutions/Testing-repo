
# ==================== Imports ====================
print("Second - Started Auto code")
import os
import sys
import json
import socket
import logging
import subprocess
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
from googledriver import download_folder
import queue

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

EXCEPTION_FOLDERS = {"excellFiles", "icons", "logData", "textDocuments"}
DELETE_FOLDERS_FILES = [
    "excellFiles", "icons", "logData", "textDocuments",
    "autoUpdate.py", "autoUpdate.pyc",
    "pythonCode.py", "pythonCode.pyc",
    "systemConfig.ini", "iconList.ini"
]

# ==================== Logging Setup ====================
console = Console()
install(show_locals=True)

def setup_logger(name, log_file):
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    if logger.handlers:
        return logger
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
with open(os.path.join(PROJECT_ROOT, "data.json"), "r") as f:
    input_data = json.load(f)
data = input_data

# ==================== Config Parsing ====================
CONFIG_KEYS = [
    "SBN_ID", "CLIENT_NAME", "LOCATION", "CITY", "STATE", "COUNTRY",
    "RENTAL_OR_SUBSCRIPTION", "RENEWAL_DATE", "COBRANDING",
    "EVENT_TYPE", "EVENT_NUMBER", "QR_ENABLED", "QR_EMAILID",
    "QR_PAYMENT_URL", "SHEETID"
]

config_values = {}

def load_config():
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

QUIT = True

# ==================== NETWORK ====================
def check_internet():
    try:
        socket.setdefaulttimeout(5)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("8.8.8.8", 53))
        return True
    except Exception as e:
        systemLogger.debug(f"Internet Check Failed: {e}")
        return False

# ==================== UI Setup ====================
WINDOW_WIDTH, WINDOW_HEIGHT = 800, 480
whiteColor = "#ffffff"
textColor = "#652828"

window = Tk()
window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
window.resizable(False, False)
window.overrideredirect(True)

content_page = Frame(window, width=WINDOW_WIDTH, height=WINDOW_HEIGHT, bg=whiteColor)

infoimages = {}
infoTexts = {}
ui_queue = queue.Queue()

try:
    infoimages["getting_information"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "Downloading_red.png")))
    infoimages["wifi_error"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "Downloading_NoWifiErrorIcon.png")))
    infoimages["rand_error"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "Downloading_RandonErrorIcon.png")))
    infoimages["usage_update"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "Downloading_UsageDataUpdateIcon.png")))
    infoimages["content_update"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "Downloading_ContentUpdateIcon.png")))
    infoimages["update_done"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "Downloading_UpdateDone.png")))

    infoimages["quit_app"] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "error_quit.png")))
except Exception as e:
    systemLogger.warning(f"Image load failed: {e}")

infoTexts["getting_info"] = "Getting content information..."
infoTexts["wifi_connected"] = "Internet Connected!"
infoTexts["wifi_error"] = "No Internet"
infoTexts["rand_error"] = "Random Error - Unable to Update"
infoTexts["update_done"] = "Update Done Sucessfully!"

update_content_canvas = None
comment_text_id = None

# ==================== UI Helpers ====================
def request_quit():
    global QUIT
    QUIT = False
    try:
        if SYSTEM_PLATFORM:
            os.system('vcgencmd display_power 0')
    finally:
        systemLogger.info("Application exiting...")
        window.destroy()

def raise_frame(frame):
    if SYSTEM_PLATFORM:
        os.system('vcgencmd display_power 1')
    frame.tkraise()
    frame.pack()

# ==================== Content Update Logic ====================
def updateDownload_File(downloaderRun):
    try:
        while downloaderRun.is_set():
            try:
                folders = glob.glob(os.path.join(PROJECT_ROOT, '*/'))
                if not folders:
                    sleep(1)
                    continue
                latest_folder = max(folders, key=os.path.getmtime)
                last_folder_name = os.path.basename(os.path.normpath(latest_folder))
                if last_folder_name not in EXCEPTION_FOLDERS:
                    file_count = len(os.listdir(latest_folder))
                    info_text = f"Downloading file number {file_count} from {last_folder_name}"

                    systemLogger.info(info_text)
                    ui_queue.put(info_text)
                else:
                    systemLogger.info("Monitoring download folders - No update yet.")
            except Exception as e:
                systemLogger.warning(f"Folder monitoring error: {e}")
                pass
            sleep(1)
    except Exception as e:
        systemLogger.error(f"Critical error in updateDownload_File: {e}")
        pass

def fileUpdates(CUMULATIVE_EXCEL_FILE_NAME):
    try:
        cumulative_sheets = pd.read_excel(CUMULATIVE_EXCEL_FILE_NAME, sheet_name=None)
        sheet1 = cumulative_sheets.get("Sheet1")
        sheet2 = cumulative_sheets.get("Sheet2")
        if sheet1 is None or sheet2 is None:
            systemLogger.error("Missing required sheets in cumulative Excel.")
            return
        
        missingFilePath = []
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
                thread = threading.Thread(target=updateDownload_File, args=(downloaderRun,), daemon=True)
                thread.start()
                try:
                    download_folder(version_url, download_dir)
                finally:
                    downloaderRun.clear()
                    thread.join(timeout=5)
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
        if missingFilePath:
            with pd.ExcelWriter(CUMULATIVE_EXCEL_FILE_NAME, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
                pd.DataFrame(missingFilePath, columns=["Downloaded Files"]).to_excel(writer, sheet_name="DownloadedPath", index=False)
    except Exception as e:
        systemLogger.error(f"Error during file check: {e}")
        ui_queue.put("Error during file update")
        ui_queue.put("__QUIT__")

    upload_file("CONTENT_CROSSCHECK_FILE")

def fileCheck(CUMULATIVE_EXCEL_FILE_NAME, SYSTEM_EXCEL_FILE_NAME):
    try:
        system_df = pd.read_excel(SYSTEM_EXCEL_FILE_NAME)
        cumulative_sheets = pd.read_excel(CUMULATIVE_EXCEL_FILE_NAME, sheet_name=None)
        sheet1 = cumulative_sheets.get("Sheet1")
        if sheet1 is None:
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
                try:
                    os.remove(norm_path)
                    deleted_paths.append(path)
                except Exception as e:
                    pass
        if deleted_paths:
            with pd.ExcelWriter(CUMULATIVE_EXCEL_FILE_NAME, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
                pd.DataFrame(deleted_paths, columns=["Deleted Files"]).to_excel(writer, sheet_name="DeletedPath", index=False)
    except Exception as e:
        ui_queue.put("Error during file check")
        ui_queue.put("__QUIT__")
    fileUpdates(CIMMULATIVE_EXCEL_FILE_NAME)

def contentUpdate(excelFileID):
    try:
        ui_queue.put("Updating content")
        excel_url = f"https://drive.google.com/uc?id={excelFileID}"
        gdown.download(url=excel_url, 
                       output=CIMMULATIVE_EXCEL_FILE_NAME, 
                       quiet=False, fuzzy=True, use_cookies=False)
        directory = os.path.join(PROJECT_ROOT, "textDocuments")
        file_list = glob.glob(directory + "/**/*.jpg", recursive=True)
        file_names = [os.path.basename(f) for f in file_list]
        df = pd.DataFrame({"Path": file_list, "fileName": file_names})
        df.to_excel(SYSTEM_EXCEL_FILE_NAME, index=False)
        fileCheck(CIMMULATIVE_EXCEL_FILE_NAME, SYSTEM_EXCEL_FILE_NAME)
    except Exception as e:
        ui_queue.put("Content update failed")
        ui_queue.put("__QUIT__")
# ==================== Cloud Communication ====================
def post_with_retry(url, *, params=None, data=None, timeout=15, retries=2, backoff=2.0):
    for attempt in range(retries + 1):
        try:
            return requests.post(url, params=params, data=data, timeout=timeout)
        except Exception as e:
            if attempt >= retries:
                raise
            systemLogger.warning(f"POST failed (attempt {attempt+1}/{retries+1}): {e}")
            sleep(backoff ** attempt)

def upload_file(file_type):
    try:
        now = datetime.now()
        prev_month = now - relativedelta(months=1)
        if file_type == "USAGE_DATA_EXCEL":
            fname = f"{config_values['SBN_ID']}_{prev_month.strftime('%m')}-{prev_month.year}.xls"
        elif file_type == "USAGE_DATA_EXCEL_BACKUP":
            fname = f"{config_values['SBN_ID']}_backup.log"
        elif file_type == "CONTENT_CROSSCHECK_FILE":
            fname = "cumulativeExcelSheet.xlsx"
            #tempUploadName = f"{config_values['SBN_ID']}_cumulativeExcelSheet.xlsx"
        else:
            return
        if "USAGE" in file_type:
            fpath = os.path.join(EXCELLSHEET_ROOT, fname)
        else:
            fpath = os.path.join(PROJECT_ROOT, fname)
        if os.path.exists(fpath):
            #ui_queue.put("Sending usage file " + fname)
            systemLogger.info("Sending usage file " + fname)
            ui_queue.put("Sending usage file")
            with open(fpath, "rb") as f:
                encoded = base64.urlsafe_b64encode(f.read())
        else:
            systemLogger.warning("File Not Found... " + fname)
            ui_queue.put("File Not Found...")
            encoded = None
        params = {
            "StoryboxId": config_values["SBN_ID"], 
            "Configuration": file_type, 
            "FileName": fname
            }
        global QUIT
        if(QUIT):
            post_with_retry(SCRIPT_URL, params=params, 
                            data=encoded if encoded else None)
    except Exception as e:
        ui_queue.put("Upload error")

# ==================== Threadding ====================
def run_content_updater():
    try:
        if not check_internet():
            ui_queue.put("Internet Not Connected")
            ui_queue.put("__QUIT__")
            return
        today = int(datetime.now().strftime("%Y%m%d"))
        renewal_due = int(data[0]["RENEWAL_DATE"]) <= today

        if not renewal_due:
            if data[0].get("USAGE_DATA", "YES") == "NO":
                upload_file("USAGE_DATA_EXCEL")
                upload_file("USAGE_DATA_EXCEL_BACKUP")
            if data[0].get("CONTENT_UPDATE", "YES") == "NO":
                contentUpdate(data[0]["CONTENT_UPDATE_EXCEL_LINK"])
            
            main_content_ui(imgName=infoimages["update_done"], infoText=infoTexts["update_done"])
            sleep(2)
            ui_queue.put("__QUIT__")
    except Exception as e:
        ui_queue.put("Code update error")
        ui_queue.put("__QUIT__")

# ==================== UI Actions ====================
def _drain_ui_queue():
    global update_content_canvas, comment_text_id
    try:
        while True:
            try:
                msg = ui_queue.get_nowait()
            except queue.Empty:
                break
            if msg == "__QUIT__":
                window.after(0, window.destroy())
                continue
            if update_content_canvas is not None and comment_text_id is not None:
                try:
                    update_content_canvas.itemconfigure(comment_text_id, text=msg)
                except Exception:
                    pass
    finally:
        window.after(100, _drain_ui_queue)

def main_content_ui(imgName=None, infoText=None):
    global update_content_canvas, comment_text_id

    if imgName is None:
        imgName = infoimages.get("getting_information")
    if infoText is None:
        infoText = infoTexts.get("getting_info", "")

    if update_content_canvas is None:
        update_content_canvas = Canvas(content_page, bg=whiteColor, 
                                       width=WINDOW_WIDTH, height=WINDOW_HEIGHT, 
                                       bd=0, highlightthickness=0)
        update_content_canvas.place(x=0, y=0)

        update_content_canvas.create_image(400, 220, 
                                           image=imgName, anchor=CENTER)

        comment_text_id = update_content_canvas.create_text((400, 45), 
                                                            text=infoText, font="Candara 16 bold", fill=textColor)
        
        Button(update_content_canvas, image = infoimages["quit_app"], 
               bg="white", borderwidth=0, command=request_quit).place(x=795, y=0)
        
        if SYSTEM_PLATFORM:
            os.system('vcgencmd display_power 1')

        raise_frame(content_page)
        window.after(100, _drain_ui_queue)
    else:
        update_content_canvas.itemconfigure(comment_text_id, text=infoText)
        update_content_canvas.create_image(400, 220, image=imgName, anchor=CENTER)
        #raise_frame(content_page)

# ==================== Start ====================
if __name__ == "__main__":
    main_content_ui()
    threading.Thread(target=run_content_updater, daemon=True).start()
    #window.after(1, run_content_updater)
    window.mainloop()
    
    if SYSTEM_PLATFORM:
        os.system('vcgencmd display_power 0')

    MAIN_UI_CODE = "interfacecode.py"

    CMD = "sudo python3 "
    if(SYSTEM_PLATFORM):
        os.system(CMD + os.path.join(PROJECT_ROOT, MAIN_UI_CODE))
    else:
        subprocess.run(["python", os.path.join(PROJECT_ROOT, MAIN_UI_CODE)])
