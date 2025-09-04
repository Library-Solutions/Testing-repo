import os
import json
import socket
import shutil
import logging
import requests
import threading
from datetime import datetime
from time import sleep
import subprocess
from urllib.request import urlopen
from configparser import ConfigParser
from tkinter import Tk, Canvas, Frame, Button, CENTER
from PIL import Image, ImageTk

from rich.console import Console
from rich.logging import RichHandler
from rich.traceback import install
import pyfiglet

import ntplib
from datetime import datetime, timezone
import subprocess
import ctypes

stop_event = threading.Event()

import sys
sys.setrecursionlimit(5000)
if(sys.platform == "win32" or sys.platform == "darwin"):
    SYSTEM_PLATFORM = False
else:
    SYSTEM_PLATFORM = True

if(SYSTEM_PLATFORM):
    import neopixel
    import board
    import digitalio

    os.system('vcgencmd display_power 0')
    pixels = neopixel.NeoPixel(board.D21, 60)
    paperSensor = digitalio.DigitalInOut(board.D5)
    paperSensor.direction = digitalio.Direction.INPUT

# ==================== PATH SETUP ====================
PROJECT_ROOT = os.path.dirname(os.path.realpath(__file__))
ICONS_ROOT = os.path.join(PROJECT_ROOT, 'icons')
LOG_FILE = os.path.join(PROJECT_ROOT, 'logData', 'system.log')
#LOGDATA_ROOT = os.path.join(PROJECT_ROOT, 'logData')
CONFIG_FILE = os.path.join(PROJECT_ROOT, 'systemConfig.ini')

# Log files
SYSTEM_LOG_FILE = os.path.join(LOG_FILE)

# ==================== LOGGER SETUP ====================
console = Console()
install(show_locals=True)
formatter = logging.Formatter("%(asctime)s :: %(levelname)-8s :: %(message)s", "%Y-%m-%d %H:%M:%S")

def setup_logger(log_file, logger_name):                                                   # Function to set up logging
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.DEBUG)
    console_handler = RichHandler(console=console)                                          # RichHandler for console output
    console_handler.setFormatter(formatter)
    file_handler = logging.FileHandler(log_file, encoding="utf-8", mode="a")                # FileHandler for writing log data to a file
    file_handler.setFormatter(formatter)
    logger.addHandler(console_handler)                                                      # Add handlers to the logger
    logger.addHandler(file_handler)
    return logger

systemLogger = setup_logger(SYSTEM_LOG_FILE, "systemLogger")                                # Set up system and usage loggers

# ==================== UI SETUP ====================
WINDOW_WIDTH, WINDOW_HEIGHT = 800, 480
window = Tk()
window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
window.resizable(False, False)
window.overrideredirect(True)

update_page = Frame(window, width=WINDOW_WIDTH, height=WINDOW_HEIGHT, bg="#ffffff")
update_canvas = None
comment_text_id = None
image_id = None

# ==================== ASCII ART ====================
def display_ascii_art(text):
    ascii_art = pyfiglet.figlet_format(text)
    console.print(ascii_art)

display_ascii_art("Story Box!\nUpdater")

# ==================== CONFIG ====================
CONFIG_KEYS = [
    "SBN_ID", "CLIENT_NAME", "LOCATION", "CITY", "STATE", "COUNTRY",
    "RENTAL_OR_SUBSCRIPTION", "RENEWAL_DATE", "COBRANDING",
    "EVENT_TYPE", "EVENT_NUMBER", "QR_ENABLED", "QR_EMAILID",
    "QR_PAYMENT_URL", "SHEETID"
]

config_values = {}
config = ConfigParser()

def loadConfigFileData():
    config.read(CONFIG_FILE)
    storybox_info = config['storyboxinfo']

    for key in CONFIG_KEYS:
        value = storybox_info.get(key)
        config_values[key] = value
        systemLogger.info(f"{key}: {value}")
        #print(f"{key}: {value}")
    config_values["EVENT_NUMBER"] = int(config_values.get("EVENT_NUMBER", 0))
    config_values["RENEWAL_DATE"] = int(config_values.get("RENEWAL_DATE", 0))

loadConfigFileData()

SCRIPT_URL = f"https://script.google.com/macros/s/{config_values['SHEETID']}/exec"

# ==================== NETWORK ====================
def check_internet():
    try:
        socket.setdefaulttimeout(10)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect(("8.8.8.8", 53))
        systemLogger.debug("Internet Connected")
        return True
    except Exception as e:
        systemLogger.warning(f"Internet Check Failed: {e}")
        return False

# ==================== SHEET FETCH ====================
def fetchSheetUpdateData():
    try:
        response = requests.get(SCRIPT_URL, params={
            'StoryboxId': config_values["SBN_ID"],
            'Configuration': "GET_SHEET_DATA",
        })
        return json.loads(response.text)['data']
    except Exception as e:
        systemLogger.error(f"Sheet fetch failed: {e}")
        return None

# ==================== CONFIG UPDATE ====================
def updateSystemConfigFile(data):
    try:
        for key in CONFIG_KEYS:
            if key != "SHEETID":
                config['storyboxinfo'][key] = str(data[0][key])
        with open(CONFIG_FILE, 'w') as configfile:
            config.write(configfile)
        systemLogger.info("System config updated")
    except Exception as e:
        systemLogger.error(f"Config update failed: {e}")

# ==================== UI =================================
def updateScreen(image="Downloading_red.png", message="Getting information..."):
    global update_canvas, comment_text_id, image_id

    try:
        update_canvas = Canvas(update_page, bg="#ffffff", width=WINDOW_WIDTH, height=WINDOW_HEIGHT, bd=0, highlightthickness=0)
        update_canvas.place(x=0, y=0)

        img = Image.open(os.path.join(ICONS_ROOT, image))
        update_img = ImageTk.PhotoImage(img)
        update_canvas.update_img = update_img  # Prevent garbage collection
        image_id = update_canvas.create_image(400, 220, image=update_img, anchor=CENTER)

        comment_text_id = update_canvas.create_text((400, 45), text=message, font="Candara 16 bold", fill="#652828")

        quit_img = Image.open(os.path.join(ICONS_ROOT, "error_quit.png"))
        quit_button_img = ImageTk.PhotoImage(quit_img)
        Button(update_canvas, image=quit_button_img, bg='white', borderwidth=0, command=quitApp).place(x=795, y=0)
        update_canvas.quit_button_img = quit_button_img

        raise_frame(update_page)
    except Exception as e:
        systemLogger.error(f"UI Error: {e}")

def updateMessage_UI(text, image=None):
    def _update():
        if image:
            img = Image.open(os.path.join(ICONS_ROOT, image))
            img_tk = ImageTk.PhotoImage(img)
            update_canvas.update_img = img_tk
            update_canvas.itemconfig(image_id, image=img_tk)
        update_canvas.itemconfigure(comment_text_id, text=text)
    window.after(0, _update)

def raise_frame(frame):
    os.system('vcgencmd display_power 1')
    frame.tkraise()
    frame.pack()

def updateSucessInfoCloudDashboard(value):
    if SCRIPT_URL:
        try:
            requests.get(SCRIPT_URL, params={
                'StoryboxId': config_values.get('SBN_ID', 'UNKNOWN'),
                'Configuration': value,
            })
            systemLogger.info("Notified cloud of code update completion.")
        except Exception as notify_error:
            systemLogger.warning(f"Notification failed: {str(notify_error)}")

# Function to check internet connectivity
def is_connected(host="github.com", port=443, timeout=5):
    try:
        socket.create_connection((host, port), timeout=timeout)
        return True
    except OSError:
        return False

def downloadCodeUpdate(URL):
    try:
        # Wait for network (up to 60 seconds)
        for i in range(12):
            if is_connected():
                systemLogger.info("Network is up. Proceeding with clone...")
                break
            else:
                systemLogger.warning(f"Network not ready, retrying in 5 seconds... ({i+1}/12)")
                sleep(5)
        else:
            systemLogger.error("Failed to detect network after 60 seconds. Exiting.")
            return

        # Extract folder name from URL
        # Prepare folder names
        folder_name = URL.split("/")[-1].replace(".git", "")
        current_dir = os.path.dirname(os.path.abspath(__file__))
        clone_path = os.path.join(current_dir, folder_name)
        
        # Remove existing repo folder
        if os.path.exists(clone_path):
            shutil.rmtree(clone_path)

        # Clone the repo
        subprocess.run(["git", "clone", URL, clone_path], check=True)
        systemLogger.info(f"Cloned repository to: {clone_path}")

        # Copy files from clone (skip .git folder)
        for file_name in os.listdir(clone_path):
            src_path = os.path.join(clone_path, file_name)
            dst_path = os.path.join(PROJECT_ROOT, file_name)

            if file_name != ".git":
                if os.path.isdir(src_path):
                    if os.path.exists(dst_path):
                        shutil.rmtree(dst_path)
                    shutil.copytree(src_path, dst_path)
                else:
                    shutil.copy2(src_path, dst_path)
                systemLogger.info(f"Copied {file_name} to project root.")
        
        # Optional cleanup
        if SYSTEM_PLATFORM:
            shutil.rmtree(clone_path)

        updateSucessInfoCloudDashboard("CODE_UPDATE_DONE")

    except Exception as e:
        updateMessage_UI(f"downloadCodeUdpate error", "Downloading_RandonErrorIcon.png")
        systemLogger.error(f"downloadCodeUdpate error: {str(e)}")

# ==================== QUIT ====================
def quitApp():
    os.system('vcgencmd display_power 0')
    systemLogger.info("Application exiting...")
    window.destroy()
    sys.exit()

def loadMainUI_PythonFile():
    os.system('vcgencmd display_power 0')
    tempPythonFilePath = ""
    if(SYSTEM_PLATFORM):
        tempPythonFilePath = 'sudo '
    tempPythonFilePath = tempPythonFilePath + 'python3 ' + str(PROJECT_ROOT) + '/pythonCode.py'
    if(os.path.exists(tempPythonFilePath)):
        #os.system(tempPythonFilePath)
        subprocess.Popen(['sudo' ,'python3', 'pythonCode.py'])
    else:
        subprocess.Popen(['sudo' ,'python3', 'pythonCode.pyc'])
    window.destroy()
    sys.exit() 


def loadautoUpdate_PythonFile(data):
    os.system('vcgencmd display_power 0')
 
    tempPythonFilePath = ""
    if(SYSTEM_PLATFORM):
        tempPythonFilePath = 'sudo '
        tempPythonFilePath = tempPythonFilePath + 'python3 ' + os.path.join(PROJECT_ROOT , 'auto.py')
        try:
            with open("data.json", "w") as f:
                json.dump(data, f)

            stop_event.set()
            #os.system(tempPythonFilePath)
            subprocess.Popen(['sudo' ,'python3', 'auto.py'])
            print("New script launched.")
            #window.destroy()
            stop_event.set()
            sys.exit() 
        except Exception as e:
            print(f"[Error] Failed to launch script: {e}")

# Now safely close the current Tkinter window
        #window.destroy()
        #print("Closed")
        #if(os.path.exists(tempPythonFilePath)):
        #window.destroy()
        #print(os.system(tempPythonFilePath + " " + str(data)))
        #print("Closed")
        #else:
        #    os.system(tempPythonFilePath+"c" + " " + str(data))

# ==================== MAIN UPDATE LOGIC ====================
def run_update():
    updateMessage_UI("Checking Internet...", "Downloading_red.png")
    attempts = 0
    while attempts < 100 and not check_internet():
        updateMessage_UI(f"Waiting for Internet... {attempts+1}", "Downloading_red.png")
        sleep(1)
        attempts += 1

    if not check_internet():
        updateMessage_UI("Internet Not Connected", "Downloading_NoWifiErrorIcon.png")
        #Get time from the RTC no Internet Connected
        if (SYSTEM_PLATFORM):
            #subprocess.run(['sudo', 'hwclock', '--hctosys'], check=True)
            systemLogger.debug("Unable to connect to Internet - Updating from the RTC Clock.")
        sleep(2)

        if(SYSTEM_PLATFORM):
            systemLogger.info("Loading the main Python UI File as no Internet")
            loadMainUI_PythonFile()
        else:
            systemLogger.debug("Dev mode - Code Running in windows - Loading the main Python UI File as no Internet")

    #-------------------------------------------------------------------------------# 
    #Get time from the Internet
    """
    updateMessage_UI("Updating System Date & Time", "Downloading_red.png")
    ntp_time = get_ntp_time()
    if not ntp_time:
        updateMessage_UI("Unable to Update Date & Time", "Downloading_NoWifiErrorIcon.png")
        if (SYSTEM_PLATFORM):
            subprocess.run(['sudo', 'hwclock', '--hctosys'], check=True)
            systemLogger.debug("Unable to set the time from the Internet - Updating from the RTC Clock.")
        sleep(2)
    if (SYSTEM_PLATFORM):
        systemLogger.debug(f"[Info] NTP Time (UTC): {ntp_time}")
        set_system_time_linux(ntp_time)
        sleep(1)  # wait a moment to ensure time is set before syncing RTC
        sync_rtc_linux()
    #-------------------------------------------------------------------------------#
    """
    data = fetchSheetUpdateData()
    #print(data)
    if not data:
        updateMessage_UI("Failed to fetch cloud data", "Downloading_RandonErrorIcon.png")
        sleep(1)
        
        if(SYSTEM_PLATFORM):
            systemLogger.info("Loading the main Python UI File as no data received from Cloud")
            loadMainUI_PythonFile()
        else:
            systemLogger.debug("Dev mode - Code Running in windows - Loading the main Python UI File as no data received from Cloud")

    today = int(datetime.now().strftime("%Y%m%d"))

    if "INFO:" in data:
        updateMessage_UI("No Update Available", "Downloading_UpdateDone.png")
        if(SYSTEM_PLATFORM):
            loadMainUI_PythonFile()
        else:
            systemLogger.debug("Dev mode - Code Running in windows - Loading the main Python UI File - No Update available")
    elif "ERROR:" in data:
        updateMessage_UI("Server Error", "Downloading_RandonErrorIcon.png")
    else:
        updateMessage_UI("Updating Configuration...", "Downloading_ConfigUpdateIcon.png")
        sleep(1)
        updateSystemConfigFile(data)
        #Reloading the data after the Config Update
        loadConfigFileData()

        #Update will be done only if the renewal is not due
        if((int(data[0]["RENEWAL_DATE"])) > int(today)):
            if (data[0]["CODE_UPDATE"] == "NO"):
                updateMessage_UI("Downloading Code Update...", "Downloading_CodeUpdateIcon.png")
                systemLogger.info(f"Code Update Link: {data[0]['CODE_UPDATE_LINK']}")
                downloadCodeUpdate(data[0]["CODE_UPDATE_LINK"])  # Placeholder

            if (data[0]["CONTENT_UPDATE"] == "NO") or (data[0]["USAGE_DATA"] == "NO"):
                systemLogger.info("AXZ Loading the main autoUpdate UI File as the Usage or Content need to be updated")
                loadautoUpdate_PythonFile(data)

            if(SYSTEM_PLATFORM):
                loadMainUI_PythonFile()
            else:
                systemLogger.debug("Dev mode - Code Running in windows - Loading the main Python UI File - No Update available") 

        else:
            if(SYSTEM_PLATFORM):
                systemLogger.info("Loading the main Python UI File as the renewal is due") 
                loadMainUI_PythonFile()
            else:
                systemLogger.debug("Dev mode - Code Running in windows - Loading the main Python UI File as the renewal is due")  

    #sleep(4)


#-------------------------------------------------------------------------------#
#Get time from the Internet
def get_ntp_time(server="pool.ntp.org"):
    try:
        client = ntplib.NTPClient()
        response = client.request(server, version=3)
        return datetime.fromtimestamp(response.tx_time, timezone.utc)
    except Exception as e:
        print(f"[Error] Failed to get NTP time: {e}")
        return None


def set_system_time_linux(dt):
    try:
        date_str = dt.strftime('%m%d%H%M%Y.%S')  # Format: MMDDhhmmYYYY.SS
        subprocess.run(['sudo', 'date', date_str], check=True)
        print("[Info] System time updated successfully on Linux.")
    except subprocess.CalledProcessError as e:
        print(f"[Error] Failed to set time on Linux: {e}")


def sync_rtc_linux():
    try:
        subprocess.run(['sudo', 'hwclock', '--systohc'], check=True)
        print("[Info] RTC clock synchronized with system time on Linux.")
    except subprocess.CalledProcessError as e:
        print(f"[Error] Failed to sync RTC: {e}")

 #-------------------------------------------------------------------------------#


# ==================== START ====================
if __name__ == "__main__":
    updateScreen()
    threading.Thread(target=run_update, daemon=True).start()
    window.mainloop()
