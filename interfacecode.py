#Code Updated On 26-03-2024
#VersioControl = "Version: 3.0\nMulti-Language With Pay Per Use Support"
#print(VersioControl)

from tkinter import *
from PIL import Image, ImageTk

import socket
import sqlite3
import threading
import logging

import xlrd
import xlwt
from xlutils.copy import copy
import socket

import urllib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

from webdriver_manager.chrome import ChromeDriverManager

import os
import sys
from datetime import *  
from time import gmtime,strftime,sleep
sys.setrecursionlimit(5000)

import json
from datetime import *  
from dateutil.relativedelta import relativedelta

if(sys.platform == "win32" or sys.platform == "darwin"):
    systemPlatform = False
else:
    systemPlatform = True

if(systemPlatform):
    from escpos.printer import Usb
    import neopixel
    import board
    import digitalio

    os.system('vcgencmd display_power 0')

    pixels = neopixel.NeoPixel(board.D21, 60)

    paperSensor = digitalio.DigitalInOut(board.D5)
    paperSensor.direction = digitalio.Direction.INPUT

browser = ""

import usb.core
import usb.util

# Replace with your printer's Vendor ID and Product ID
VENDOR_ID = 0x1504
PRODUCT_ID = 0x0103

#device = usb.core.find(idVendor=VENDOR_ID, idProduct=PRODUCT_ID)

try:
    device = usb.core.find(idVendor=VENDOR_ID, idProduct=PRODUCT_ID)
    p = Usb(0x1504,0x0103,in_ep=0x81, out_ep=0x02)
except Exception as e:
    p = "Error"

LANGUAGE_SELECTION = ""
SUB_ONE_SELECTION = ""
SUB_TWO_SELECTION = ""
SUB_THREE_SELECTION = ""

returnIndexValue = 0
storedDate = 0
loggedData = ""

PROJECT_ROOT = os.path.dirname(os.path.realpath(__file__))

EXCELLSHEET_ROOT = os.path.join(PROJECT_ROOT , "excellFiles")
CONFIG_FILE = os.path.join(PROJECT_ROOT, 'systemConfig.ini')
DOCUMENTS_ROOT = os.path.join(PROJECT_ROOT, 'textDocuments')
LOGDATA_ROOT = os.path.join(PROJECT_ROOT, 'logData')
ICONS_ROOT = os.path.join(PROJECT_ROOT, 'icons')


DELETE_FOLDERS_FILES = ["excellFiles","icons", "logData", "textDocuments",
                        "systemConfig.ini","iconList.ini",
                        "autocodeupdate.py", "autocodeupdate.pyc",
                        "contentupdatecode.py", "contentupdatecode.pyc",
                        "interfacecode.py", "interfacecode.pyc",
                        "codeUpdate.py", "auto.py",
                        "pythonCode.py", "two_contentupdatecode.py"]

tempFilesName = os.listdir(PROJECT_ROOT)
print(tempFilesName)
if(systemPlatform):
    for i in range(len(tempFilesName)):
        if(tempFilesName[i] not in DELETE_FOLDERS_FILES):
            os.system("sudo rm -r "+os.path.join(str(PROJECT_ROOT), str(tempFilesName[i])))

windowWidth = 800
windowHeight = 480

buttonWidth = 175
buttonHeight = 195

buttonAlign_Y = 260-110
buttonAlign_X = 110-90

backButtonWidthHeight = 80

whiteColor = "#ffffff"

window = Tk()
window.geometry("800x480")
window.resizable(False, False)
window.overrideredirect(True)

import configparser

from rich.logging import RichHandler
from rich.console import Console
console = Console()                             # Console setup for rich output
from rich.text import Text
from rich.progress import track
from rich.traceback import install
install(show_locals=True)                       # Initialize rich traceback for displaying local variables in stack traces

import logging
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

import pyfiglet                                 # ASCII Art Display
def display_ascii_art(text):
    ascii_art = pyfiglet.figlet_format(text)
    console.print(ascii_art)

display_ascii_art("Story Box!")                 # Display ASCII Art for Story Box



# Read the system config from the config.ini
SYSTEM_ICONSLIST_READER = configparser.ConfigParser()
SYSTEM_ICONSLIST_READER.read(os.path.join(PROJECT_ROOT, 'iconList.ini'))
SYSTEM_ICONSLIST_CONFIG = SYSTEM_ICONSLIST_READER['storyboxList']


replacements = {' ' : '', '[' : '', ']': '', '"': '' }
#LANGUAGE_LIST = ["telugu", "english", "hindi"]
#LANGUAGE_LIST =  (SYSTEM_ICONSLIST_CONFIG.get("LANGUAGE_LIST")).replace([' ','"'], ['','']).split(",")
LANGUAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("LANGUAGE_LIST")).split(",")
print((LANGUAGE_LIST))

##TELUGU LISTS
TELUGUDASH_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUDASH_LIST")).split(",")
TELUGUGENRE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUGENRE_LIST")).split(",")
TELUGUAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUAGE_LIST")).split(",")

TELUGUBIP_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUBIP_LIST")).split(",")
TELUGUFACTS_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUFACTS_LIST")).split(",")

TELUGUPUZZLES_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUPUZZLES_LIST")).split(",")
TELUGUPUZZLESAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("TELUGUPUZZLESAGE_LIST")).split(",")


##HINDI LISTS
HINDIDASH_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIDASH_LIST")).split(",")
HINDIGENRE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIGENRE_LIST")).split(",")
HINDIAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIAGE_LIST")).split(",")

HINDIBIP_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIBIP_LIST")).split(",")
HINDIFACTS_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIFACTS_LIST")).split(",")

HINDIPUZZLES_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIPUZZLES_LIST")).split(",")
HINDIPUZZLESAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("HINDIPUZZLESAGE_LIST")).split(",")
"""
HINDIDASH_LIST = ["hindiDash_puzzles"]
HINDIPUZZLES_LIST = ["hindiPuzzles_crossWord"]
HINDIPUZZLESAGE_LIST = ["hindiPuzzlesAge_easy", "hindiPuzzlesAge_medium", "hindiPuzzlesAge_high"]
"""
##ENGISH LIST
ENGLISHDASH_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHDASH_LIST")).split(",")
ENGLISHGENRE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHGENRE_LIST")).split(",")
ENGLISHAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHAGE_LIST")).split(",")

ENGLISHBIP_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHBIP_LIST")).split(",")
ENGLISHFACTS_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHFACTS_LIST")).split(",")

ENGLISHPUZZLES_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHPUZZLES_LIST")).split(",")
ENGLISHPUZZLESAGE_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("ENGLISHPUZZLESAGE_LIST")).split(",")

EVENT_LIST = "".join(replacements.get(char, char) for char in SYSTEM_ICONSLIST_CONFIG.get("EVENT_LIST")).split(",")


# Read the system config from the config.ini
SYSTEM_CONFIG_READER = configparser.ConfigParser()
SYSTEM_CONFIG_READER.read(os.path.join(PROJECT_ROOT, 'systemConfig.ini'))
STORYBOX_INFO_CONFIG = SYSTEM_CONFIG_READER['storyboxinfo']

CONFIG_KEYS = [
    "SBN_ID", "CLIENT_NAME", "LOCATION", "CITY", "STATE", "COUNTRY",
    "RENTAL_OR_SUBSCRIPTION", "RENEWAL_DATE", "COBRANDING",
    "EVENT_TYPE", "EVENT_NUMBER", "QR_ENABLED", "QR_EMAILID",
    "QR_PAYMENT_URL", "SHEETID"
]

config_values = {}

def load_config():
    config = configparser.ConfigParser()
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

reminderDays = 7

sheetID = str(config_values["SHEETID"])
URL = config_values["QR_PAYMENT_URL"]
qrEnabled = str(config_values["QR_ENABLED"])
qrMail = str(config_values["QR_EMAILID"])
qrLocation = str(config_values["CLIENT_NAME"] +" "+ config_values["LOCATION"])

sender_email = "storybox.sender@gmail.com"
receiver_email = "storybox.reciver@gmail.com"

USAGE_LOG_FILE = os.path.join(EXCELLSHEET_ROOT, config_values["SBN_ID"]+'_backup.log')
SYSTEM_LOG_FILE = os.path.join(LOGDATA_ROOT, 'system.log')

if not(os.path.exists(USAGE_LOG_FILE)):
    usageLogger = setup_logger(USAGE_LOG_FILE, "usageLogger")
    usageLogger.info(config_values["SBN_ID"] + " - Started")
else:
    usageLogger = setup_logger(USAGE_LOG_FILE, "usageLogger")

if not(os.path.exists(SYSTEM_LOG_FILE)):
    systemLogger = setup_logger(SYSTEM_LOG_FILE, "systemLogger")                                # Set up system and usage loggers
    systemLogger.info(config_values["SBN_ID"] + " - Started")
else:
    systemLogger = setup_logger(SYSTEM_LOG_FILE, "systemLogger")

FOLDERS_ACCESS = ["icons", "textDocuments", "logData", "excellFiles"]

if(systemPlatform):
    for i in range(len(FOLDERS_ACCESS)):
        if os.path.exists(os.path.join(PROJECT_ROOT,FOLDERS_ACCESS[i])):
            print(os.path.join(PROJECT_ROOT,FOLDERS_ACCESS[i]))
            os.system("sudo chmod -R 777 "+os.path.join(PROJECT_ROOT,FOLDERS_ACCESS[i]))

main_page = Frame(window,
            width = windowWidth,
            height = windowHeight,
            bg = whiteColor)

def raise_frame(frame):
    frame.tkraise()
    frame.pack()

def lightsColor(errorColorStatus):
    if(systemPlatform):
        if(errorColorStatus == "White"):
            pixels.fill((255, 255, 100))
        elif(errorColorStatus == "Orange"):
            pixels.fill((255, 20, 0))
        elif(errorColorStatus == "Setup"):
            pixels.fill((0, 255, 50))

def quitButton():
    global BROWSER
    if(browser != ""):
        browser.close()
        browser.quit()

    lightsColor("Setup")
    systemLogger.debug("Application Closed By The User")

    if(systemPlatform):
        os.system('vcgencmd display_power 0')
        os.system("sudo systemctl stop app.service")
    
    window.destroy()
    sys.exit()

######################################################
##*****************    LOGGER SETUP    ************###


##*************************************************##
##*******  BOOK IN PARTS db File Creation  ********##

def getCurrentDate():
    Current_Date_Formatted = datetime.today().strftime ('%Y%m%d')
    currentDate = Current_Date_Formatted
    return currentDate

def getPreviousDate():
    Previous_Date = datetime.today() - timedelta(days=1)
    PreviousDate = Previous_Date.strftime ('%Y%m%d') # format the date to ddmmyyyy
    return PreviousDate

def create_BookDataBase(LOG_FILE_ROOT, ListofIndex):
    try:
        #print("book is created", ListofIndex)
        connect = sqlite3.connect(LOG_FILE_ROOT)
        connect.execute('''CREATE TABLE SEQUENCEBOOK
                    (ID INT PRIMARY KEY     NOT NULL,
                    NAME           TEXT    NOT NULL,
                    DATE           INT    NOT NULL,
                    INDEX_NO            INT     NOT NULL,
                    NO_OF_PRINTS            INT     NOT NULL,
                    EMAIL_SENT            INT     NOT NULL);''')
        connect.commit()
            
        serialNumber = 1
        indexNumber = 0
        for x in ListofIndex:
            query = "INSERT INTO SEQUENCEBOOK (ID,NAME,DATE,INDEX_NO,NO_OF_PRINTS,EMAIL_SENT) VALUES (?,?,?,?,?,?)"
            today = getCurrentDate()
            connect.execute(query, (serialNumber,x,today,indexNumber,indexNumber,indexNumber))
            serialNumber = serialNumber+1
        connect.commit()
        connect.close()
    except Exception as error:
        systemLogger.debug("Unabe To Create create_BookDataBase Log Table - {str(error)}")

def getSequinceInfo_BookDatabase(LOG_FILE_ROOT, bookName):
    try:
        #print(LOG_FILE_ROOT)
        #print( bookName)
        connect = sqlite3.connect(LOG_FILE_ROOT)
        cursor = connect.cursor()

        query = """select * from SEQUENCEBOOK where NAME = ?"""
        cursor.execute(query, (bookName,))
        records = cursor.fetchall()

        #print(records)
            
        for row in records:
            returnIndexValue = row[3]
            storedDate = row[2]
            noOfPrints = row[4]
            emailSent = row[5]
        connect.commit()
        connect.close()
        return storedDate,returnIndexValue,noOfPrints,emailSent
    except Exception as error:
       systemLogger.debug(f"Unable to Get BOOK SEQ DATA -  {str(error)}")

def update_BookDatabase(LOG_FILE_ROOT, setValue, updateValue, bookName):
    try:
        connect = sqlite3.connect(LOG_FILE_ROOT)
        query = "UPDATE SEQUENCEBOOK set "+ setValue +" = ? where NAME = ?"
        connect.execute(query, (updateValue, bookName))
        connect.commit()
        connect.close()
    except sqlite3.Error as error:
        systemLogger.error(f"Failed to read Value data from sqlite table - {str(error)}")

def updateSequinceValue_BookDatabase(LOG_FILE_ROOT, languageOption, mainOption, bookName):
    try:
        #print(bookName)
        previousDate = getPreviousDate()
        storedDate_n, returnIndexValue, noOfPrints, emailSent = getSequinceInfo_BookDatabase(LOG_FILE_ROOT, bookName)

        path = os.path.join(DOCUMENTS_ROOT, languageOption, mainOption, bookName)
        files = os.listdir(path)
        #print(len(files))
        if ((returnIndexValue >= (len(files)-5)) and (emailSent == 0)):
            emailSent = 1
            print("email Sent", config_values["SBN_ID"])
            update_BookDatabase(LOG_FILE_ROOT,
                                "EMAIL_SENT",
                                emailSent,
                                bookName)
                
            Information = str(len(files) - returnIndexValue - 1) + " PRINTS LEFT IN " + bookName
            
            if(ping_checkInternet()):
                mailThread = threading.Thread(target=sendErrorMail, args=(Information,))
                mailThread.start()
            else:
                systemLogger.error("No Network Unable to send prints left Mail")

        if((int(previousDate) >= int(storedDate_n)) and (noOfPrints != 0)):
            if returnIndexValue < (len(files)-1):
                returnIndexValue = returnIndexValue + 1
            else:
                Information = bookName + " - Re Initialised IndexNumber to 0"

                if(ping_checkInternet()):
                    mailThread = threading.Thread(target=sendErrorMail, args=(Information,))
                    mailThread.start()
                else:
                    systemLogger.debug("No Network Unable to send Re Initialised IndexNumber Mail",'error','one')

                returnIndexValue = 0
            noOfPrints = 1          
            emailSent = 0
            update_BookDatabase(LOG_FILE_ROOT, "EMAIL_SENT", emailSent, bookName)
            update_BookDatabase(LOG_FILE_ROOT, "NO_OF_PRINTS", noOfPrints, bookName)
            update_BookDatabase(LOG_FILE_ROOT, "INDEX_NO", returnIndexValue, bookName)
        else:
            noOfPrints = noOfPrints + 1
            update_BookDatabase(LOG_FILE_ROOT, "NO_OF_PRINTS", noOfPrints, bookName)

        date = getCurrentDate()
        update_BookDatabase(LOG_FILE_ROOT, "DATE", date, bookName)
        
    except Exception as error:
        systemLogger.error(f"Unabe To Check updateSequinceValue_BookDatabase- {str(error)}")


##*************************************************##
##*************  DB File Creation  ****************##

def create_SeqDataBase(LOG_FILE_ROOT, ListofIndex):
    try:
        connect = sqlite3.connect(LOG_FILE_ROOT)
        connect.execute('''CREATE TABLE SEQUENCE
                    (ID INT PRIMARY KEY     NOT NULL,
                    NAME           TEXT    NOT NULL,
                    INDEX_NO            INT     NOT NULL);''')
        connect.commit()
        
        serialNumber = 1
        indexNumber = 0
        
        for x in ListofIndex:
            query = "INSERT INTO SEQUENCE (ID,NAME,INDEX_NO) VALUES (?,?,?)"
            connect.execute(query, (serialNumber, x, indexNumber))
            serialNumber = serialNumber+1
        connect.commit()
        connect.close()
    except Exception as error:
        systemLogger.error(f"Unabe To Create Log Table - {str(error)}")

def checkIndexValue_DataBase(LOG_FILE_ROOT, mainOption, selectedOption, DATAPATH):
    global returnIndexValue
    try:
        connect = sqlite3.connect(LOG_FILE_ROOT)
        cursor = connect.cursor()
            
        query = """select * from SEQUENCE where NAME = ?"""
        cursor.execute(query, (selectedOption,))
        records = cursor.fetchall()
        for row in records:
            returnIndexValue = row[2]
        cursor.close()
        connect.close()
            
        path = os.path.join(DOCUMENTS_ROOT , DATAPATH)
        files = os.listdir(path)
        #print(len(files))
        
        if(returnIndexValue < (len(files)-1)):
            returnIndexValue = returnIndexValue + 1
            updateSequinceValue_DataBase(LOG_FILE_ROOT, returnIndexValue, selectedOption)
        else:
            returnIndexValue = 0
            updateSequinceValue_DataBase(LOG_FILE_ROOT, returnIndexValue, selectedOption)
            
        return returnIndexValue
                
    except Exception as error:
        systemLogger.error(f"Unabe To Check checkIndexValue_DataBase Function - {selectedOption}.db Log Table - {str(error)}")

def updateSequinceValue_DataBase(LOG_FILE_ROOT, indexNo, selectedOption):
    try:
        connect = sqlite3.connect(LOG_FILE_ROOT)
        query = "UPDATE SEQUENCE set INDEX_NO = ? where NAME = ?"
        connect.execute(query, (indexNo, selectedOption))
        connect.commit()
        connect.close()
    except Exception as error:
        systemLogger.error(f"Unabe To Update {LOG_FILE_ROOT , selectedOption} Log Table {str(error)}")

def printSettings(languageOption, subOneOption, subTwoOption, subThreeOption):
    print("languageOption: ", languageOption)
    print("subOneOption: ", subOneOption)
    print("subTwoOption: ", subTwoOption)
    print("subThreeOption: ", subThreeOption)

    lightsColor("Orange")
    dataPath = ""
    indexReturn = 0
    try:
        if('Dash_bookInParts' not in subOneOption):
            if('Dash_interestingFacts' in subOneOption):
                dataPath = os.path.join(languageOption, subOneOption, subTwoOption)

                if (os.path.exists(os.path.join(LOGDATA_ROOT , subOneOption + ".db")) == False):
                    #ENGLISHFACTS_LIST
                    indexList = languageOption.upper()+"FACTS_LIST"
                    create_SeqDataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"), eval(indexList))

                indexReturn = checkIndexValue_DataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"),
                                                    subOneOption, 
                                                    subTwoOption, dataPath)
            
            elif('Dash_storyPoem' in subOneOption):
                dataPath = os.path.join(languageOption, subOneOption, subTwoOption, subThreeOption)

                if (os.path.exists(os.path.join(LOGDATA_ROOT , subOneOption + ".db")) == False):
                    tempList = []
                    tempName = languageOption.upper()
                    for i in range(len(eval(tempName + "GENRE_LIST"))):
                        for j in range(len(eval(tempName + "AGE_LIST"))):
                            tempList.append(eval(tempName + "GENRE_LIST")[i]+"-"+eval(tempName + "AGE_LIST")[j])

                    create_SeqDataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"), tempList)
                
                indexReturn = checkIndexValue_DataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"),
                                                    subOneOption, 
                                                    str(subTwoOption+"-"+subThreeOption), dataPath)

            elif('Dash_puzzles' in subOneOption):
                dataPath = os.path.join(languageOption, subOneOption, subTwoOption, subThreeOption)

                if (os.path.exists(os.path.join(LOGDATA_ROOT , subOneOption + ".db")) == False):
                    tempList = []
                    tempName = languageOption.upper()
                    for i in range(len(eval(tempName + "PUZZLES_LIST"))):
                        for j in range(len(eval(tempName + "PUZZLESAGE_LIST"))):
                            tempList.append(eval(tempName + "PUZZLES_LIST")[i]+"-"+eval(tempName + "PUZZLESAGE_LIST")[j])

                    create_SeqDataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"), tempList)
                
                indexReturn = checkIndexValue_DataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"),
                                                    subOneOption, 
                                                    str(subTwoOption+"-"+subThreeOption), dataPath)
            
            elif('event' in subOneOption):
                dataPath = os.path.join(languageOption, subOneOption, subTwoOption, subThreeOption)

                if (os.path.exists(os.path.join(LOGDATA_ROOT , languageOption + ".db")) == False):
                    create_SeqDataBase(os.path.join(LOGDATA_ROOT , languageOption + ".db"), EVENT_LIST)
                
                indexReturn = checkIndexValue_DataBase(os.path.join(LOGDATA_ROOT , languageOption + ".db"),
                                                    languageOption, 
                                                    subOneOption, dataPath)

            files = os.listdir(os.path.join(DOCUMENTS_ROOT, dataPath))
            filePath = os.path.join(DOCUMENTS_ROOT, dataPath, files[indexReturn])
        
        else:
            dataPath = os.path.join(languageOption, subOneOption, subTwoOption)
            print(dataPath)
            
            if (os.path.exists(os.path.join(LOGDATA_ROOT , subOneOption + ".db")) == False):
                tempName = languageOption.upper()
                create_BookDataBase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"), eval(tempName + "BIP_LIST"))
            
            updateSequinceValue_BookDatabase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"),
                                            languageOption,
                                            subOneOption,
                                            subTwoOption)

            storedDate, indexReturn, noOfPrints, emailSent = getSequinceInfo_BookDatabase(os.path.join(LOGDATA_ROOT , subOneOption + ".db"), subTwoOption)

            filePath = os.path.join(DOCUMENTS_ROOT, dataPath)
            print(filePath)
            book = os.listdir(filePath)
            bookName = str(book[0])
            bookName = bookName.partition("-(")

            bookName = bookName[0]+"-("+str(indexReturn)+").jpg"
            filePath = os.path.join(filePath,bookName)

        if not (os.path.exists(filePath)):
            systemLogger.debug(f"File Not Found {filePath}")
            infoScreen("", [], "error_screen", "")
        else:
            print(filePath)
            tempSelectionList = [languageOption, subOneOption, subTwoOption, subThreeOption]
            main_page.pack_forget()
            if(qrEnabled.upper() == "YES"):
                infoScreen(filePath, tempSelectionList, "zero_paymentInfo", indexReturn)
            else:
                now = datetime.now()
                tempDate = now.strftime("%Y%m%d")
            
                if(int(config_values["RENEWAL_DATE"]) < int(tempDate) and (config_values["COBRANDING"] == "YES")):
                    infoScreen(filePath, tempSelectionList, "happyReading_SchoolLogoNotAvailable", indexReturn)
                elif(int(config_values["RENEWAL_DATE"]) >= int(tempDate) and (config_values["COBRANDING"] == "YES")):
                    infoScreen(filePath, tempSelectionList, "happyReading_SchoolLogoIcon", indexReturn)

                elif(int(config_values["RENEWAL_DATE"]) < int(tempDate) and (config_values["COBRANDING"] == "NO")):
                    infoScreen(filePath, tempSelectionList, "happyReading_subscriptionRenewalScreenIcon", indexReturn)     
                elif(int(config_values["RENEWAL_DATE"]) >= int(tempDate) and (config_values["COBRANDING"] == "NO")):
                    infoScreen(filePath, tempSelectionList, "error_happyReading", indexReturn)               

    except Exception as error:
        systemLogger.error(f"printSettings - {str(error)}")
        infoScreen("", [], "error_screen", "")


def infoScreen(filePath, tempSelectionList, imageName, returnIndexValue):
    lightsColor("Orange")
    main_canvas = Canvas(main_page,
                            bg = whiteColor,
                            width = windowWidth,
                            height = windowHeight,
                            bd = 0,
                            highlightthickness = 0,
                            relief = "ridge")
    main_canvas.place(x = 0, y = 0)

    centerImage = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, imageName + ".png")))
    main_canvas.create_image(400, 220,
                                    image=centerImage,
                                    anchor=CENTER)

    quitButton_img = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "error_quit.png")))
    Button(main_canvas, bg='white', borderwidth=0, 
                image = quitButton_img, 
                command= lambda :quitButton()).place(x = 795 ,
                                                            y = 0)

    storyboxLogo = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "error_storyboxLogo.png")))
    label1 = Label(main_canvas, image=storyboxLogo, bg='white', borderwidth=0)
    label1.image = storyboxLogo
    label1.place(x=400, y = 440, anchor=CENTER)

    raise_frame(main_page)

    if("happyReading" in imageName):
        printPaperThread = threading.Thread(target=printPaper, args=(filePath, tempSelectionList, returnIndexValue, ))
        printPaperThread.start()
        
        S = threading.Timer(3, checkRentalSubscription)
        S.start()

    elif(imageName == "error_printerCommunicationError"):
        systemLogger.error(f'Printer Connection Error')
        if(ping_checkInternet()):
            sendError_MailThread = threading.Thread(target=sendErrorMail, args=('Printer Communication Error',))
            sendError_MailThread.start()
        else:
            systemLogger.error('No Internet unable to send Printer Communication Error ')

    elif(imageName == "error_paperError"):
        systemLogger.error('Paper Jam')
        if(ping_checkInternet()):
            sendError_MailThread = threading.Thread(target=sendErrorMail, args=( 'Paper Jam Error',))
            sendError_MailThread.start()
        else:
            systemLogger.error('No Internet unable to send Paper Jam mail')

    elif((imageName == "error_rentalRenewalShortly") or 
            (imageName == "error_subscriptionShortly") or 
            (imageName == "error_subscriptionRenewal")):
        #window.after(3000, iconDashboard, main_page, LANGUAGE_LIST, "first Page")
        window.after(9000, iconDashboard, main_page, LANGUAGE_LIST, "first Page")

    elif(imageName == "error_rentalRenewal"):
        systemLogger.error("Screen Locked Renewal Due")
        if(ping_checkInternet()):
            sendError_MailThread = threading.Thread(target=sendErrorMail, args=('Rental Renewal Due SYSTEM LOCKED',))
            sendError_MailThread.start()
        else:
            systemLogger.error('No Internet unable to send PRental Renewal Due SYSTEM LOCKED mail')

    elif(imageName == "zero_paymentInfo"):
        S = threading.Timer(3, checkPayment(filePath, tempSelectionList, main_canvas, centerImage, returnIndexValue))
        S.start()

    elif(imageName == "zero_paymentUnsuccessful"):
        window.after(3000, iconDashboard, main_page, LANGUAGE_LIST, "first Page")

    if(ping_checkInternet()):
        os.system('vcgencmd display_power 1')

    window.mainloop()

def sendErrorMail(Information):
    try:
        import smtplib, ssl
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        password = "uuytpuwuaacogozw"

        message = MIMEMultipart("alternative")
        message["Subject"] = "Message From StoryBox - " + config_values["SBN_ID"]
        message["From"] = sender_email
        message["To"] = receiver_email

        textBody = "Storybox ID" + config_values["SBN_ID"] + "\n" + "Information" + Information + "\n"

        #part = MIMEText(html, "html")
        part = MIMEText(textBody, "plain")
        message.attach(part)

        # Create secure connection with server and send email
        if(systemPlatform):
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL("smtp.gmail.com")
            server.login(sender_email, password)
            server.sendmail(
                sender_email, receiver_email, message.as_string()
                )
            server.quit()
        
        systemLogger.debug(f"Mail Sent {Information}")
    except  Exception as error:
        systemLogger.error(f"Error while sending an Email - {str(error)}")

######################################################
##***************SAVE USER EXCELL DATA *************##
class saveExcell_Data:        
    def saveExcell_DataCheck(self, tempTime, tempSelectionList, indexNo):
        usageLogger.debug(config_values["SBN_ID"] + "$" + tempSelectionList[0] + "$" + tempSelectionList[1] + "$" + tempSelectionList[2]+ "$" + tempSelectionList[3] + "$" + str(indexNo))
        
        excellPathFileName = os.path.join(EXCELLSHEET_ROOT, str(config_values["SBN_ID"] +"_"+tempTime[0] + ".xls"))

        if (os.path.exists(excellPathFileName) == False):
            self.StoreExcellDummyFile(tempTime)
            self.StoreExcellFile(tempTime, tempSelectionList, indexNo)
        else:
            self.StoreExcellFile(tempTime, tempSelectionList, indexNo)
            
    def StoreExcellFile(self,tempTime, tempSelectionList, indexNo):
        try:
            excellPathFileName = os.path.join(EXCELLSHEET_ROOT, str(config_values["SBN_ID"] +"_"+tempTime[0] + ".xls"))

            wb = xlrd.open_workbook(excellPathFileName)
            rowValue = wb.sheet_by_index(0)
            rowValue.cell_value(0,0)
            rowValue = rowValue.nrows
            sheetCopy = copy(xlrd.open_workbook(excellPathFileName))
            sheetCopy.get_sheet(0).write(rowValue,0,config_values["SBN_ID"])
            
            sheetCopy.get_sheet(0).write(rowValue, 1, tempTime[1])
            sheetCopy.get_sheet(0).write(rowValue, 2, tempTime[2])
            
            sheetCopy.get_sheet(0).write(rowValue, 3, tempSelectionList[0])
            sheetCopy.get_sheet(0).write(rowValue, 4, tempSelectionList[1])
            sheetCopy.get_sheet(0).write(rowValue, 5, tempSelectionList[2])
            sheetCopy.get_sheet(0).write(rowValue, 6, tempSelectionList[3])
            
            sheetCopy.get_sheet(0).write(rowValue, 7, indexNo)  
            
            sheetCopy.save(excellPathFileName)
        except Exception as error:
            systemLogger.error(f"saveExcell_Data - StoreExcellFile() - {str(error)}")
            
    def StoreExcellDummyFile(self,tempTime):
        try:
            excellPathFileName = os.path.join(EXCELLSHEET_ROOT, str(config_values["SBN_ID"] +"_"+tempTime[0] + ".xls"))

            book = xlwt.Workbook()
            sheet = book.add_sheet(tempTime[0])
            sheet.write(0,0,'Unit No')
            sheet.write(0,1,'Date')
            sheet.write(0,2,'Time')
            sheet.write(0,3,'Main Menu')
            sheet.write(0,4,'Sub Menu 1')
            sheet.write(0,5,'Sub Menu 2')
            sheet.write(0,6,'Sub Menu 3')
            sheet.write(0,7,'Index No')

            book.save(excellPathFileName)
        except Exception as error:
            systemLogger.error(f"saveExcell_Data - StoreExcellDummyFile() -{str(error)}")

######################################################

def checkPayment(filePath, tempSelectionList, payment_canvas, paymentImage, returnIndexValue):
    print("Payment Checking Started")
    global browser
    window.update()

    try:
        if(browser == ''):
            loadBrowser()            
        
        browser.get(URL)

        l = browser.find_element(By.ID ,"orderBillName")
        l.send_keys(qrLocation)

        l = browser.find_element(By.ID ,"orderBillAddress")
        l.send_keys("Gachibowli")

        l = browser.find_element(By.ID ,"orderBillZip")
        l.send_keys("500034")

        l = browser.find_element(By.ID ,"orderBillTel")
        l.send_keys("9618566556")

        l = browser.find_element(By.ID ,"orderBillEmail")
        l.send_keys(qrMail)

        browser.find_element(By.ID ,"OPTUPI").click()
        
        if(browser.find_element(By.LINK_TEXT, 'Generate QR').is_displayed()):
            browser.find_element(By.LINK_TEXT, 'Generate QR').click()
            
            sleep(2)
            images = browser.find_elements(By.TAG_NAME, 'img')
            print(len(images))
            for image in images:
                imageTemp = str(image.get_attribute('src'))
                imageSource = imageTemp.split(';')
                #print(imageSource[0])
                if(imageSource[0] == "data:image/png"):
                    #print(image.get_attribute('src'))
                    urllib.request.urlretrieve(image.get_attribute('src'), os.path.join(PROJECT_ROOT, "qrCode.png"))

                    paymentBGImage = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "zero_paymentScreen.png")))
                    payment_canvas.create_image(400, 220,
                                        image=paymentBGImage,
                                        anchor=CENTER)
                    payment_canvas.image = paymentBGImage
                    
                    paymentQRImage = ImageTk.PhotoImage(Image.open(os.path.join(PROJECT_ROOT, "qrCode.png")))
                    payment_canvas.create_image(400, 150,
                                        image=paymentQRImage,
                                        anchor=CENTER)
                    payment_canvas.image = paymentQRImage
                    
                    logo_Image = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "error_storyboxLogo.png")))
                    payment_canvas.create_image(400, 440,
                                        image=logo_Image,
                                        anchor=CENTER)
                    payment_canvas.image = logo_Image
                    
                    window.update()
                    break
        else:
            browser.find_element(By.ID ,"UPIQR").click()
            browser.find_element(By.LINK_TEXT, 'Make Payment').click()

        timeSeconds = 0
        timeOutSeconds = 30
        qr = "Not"
        
        while browser.current_url:
            window.update()
            print("Time in SECONDS: ", timeSeconds)
        
            readBody = browser.find_element(By.XPATH, "/html/body").text
            String= readBody
            firstLine = String.partition('\n')[0]
            splitString = String.split()


            if((firstLine == "QR Code for UPI Payment") and qr == "Not"):
                qr = "Yes"                

                src = browser.find_element(By.TAG_NAME, "img").get_attribute("src")
                urllib.request.urlretrieve(src, "qrCode.png")
                print("Got UPI QR")

                paymentBGImage = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "zero_paymentScreen.png")))
                payment_canvas.create_image(400, 220,
                                    image=paymentBGImage,
                                    anchor=CENTER)
                payment_canvas.image = paymentBGImage

                paymentQRImage = ImageTk.PhotoImage(Image.open("qrCode.png"))
                payment_canvas.create_image(600, 220,
                                    image=paymentQRImage,
                                    anchor=CENTER)
                payment_canvas.image = paymentQRImage

                logo_Image = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "error_storyboxLogo.png")))
                payment_canvas.create_image(400, 440,
                                        image=logo_Image,
                                         anchor=CENTER)
                payment_canvas.image = logo_Image
                window.update()

            if "Successful" in splitString:
                if(browser != ''):
                    browser.close()
                    browser.quit()
                    browser = ''
                print("payment Sucessful")
                if(int(config_values["RENEWAL_DATE"]) < int(tempDate)):
                    infoScreen(filePath, tempSelectionList, "happyReading_subscriptionRenewalScreenIcon", returnIndexValue)
                else:
                    if(os.path.exists(os.path.join(ICONS_ROOT, "happyReading_SchoolLogoIcon.png"))):
                        infoScreen(filePath, tempSelectionList, "happyReading_SchoolLogoIcon", returnIndexValue)
                    else:
                        infoScreen(filePath, tempSelectionList, "happyReading_SchoolLogoNotAvailable", returnIndexValue)
                #break

            if "Awaited" in splitString:
                if(browser != ''):
                    browser.close()
                    browser.quit()
                    browser = ''
                print("Payment Awaited Unsucessful")
                infoScreen(filePath, tempSelectionList, "zero_paymentUnsuccessful", returnIndexValue)
                #break

            if "Failed" in splitString:
                if(browser != ''):
                    browser.close()
                    browser.quit()
                    browser = ''
                print("Payment Failed Unsucessful")
                infoScreen(filePath, tempSelectionList, "zero_paymentUnsuccessful", returnIndexValue)
                #break

            elif(timeSeconds == timeOutSeconds - 5):
                paymentImage = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "zero_paymentCheck.png")))
                payment_canvas.create_image(400, 220,
                                    image=paymentImage,
                                    anchor=CENTER)
                payment_canvas.image = paymentImage
                window.update()

            elif(timeSeconds == timeOutSeconds):
                if(browser != ''):
                    browser.close()
                    browser.quit()
                    browser = ''
                print("Payment Timed Out")
                systemLogger.debug("Payment Timed Out")
                infoScreen(filePath, tempSelectionList, "zero_paymentUnsuccessful", "")

            timeSeconds = timeSeconds+1
            sleep(1)
    except Exception as error:
        if(browser != ''):
            browser.close()
            browser.quit()
            browser = ''
        print("Something else went wrong")
        systemLogger.error(f"str(error)")
        infoScreen(filePath, tempSelectionList, "zero_paymentUnsuccessful", "")


def loadBrowser():
    global browser
        
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument("--no-sandbox")
    chromeOptions.add_argument("headless")

    if(systemPlatform):
        browser_driver = Service('/usr/lib/chromium-browser/chromedriver')
        browser = webdriver.Chrome(service=browser_driver, options = chromeOptions)
                
    if(not systemPlatform):
        browser = webdriver.Chrome( ChromeDriverManager().install(), options = chromeOptions )

def getCurrent_DateTime():
    fileNameDate = datetime.today().strftime('%m-%Y')
    todayDate = datetime.today().strftime('%d-%m-%Y')
    now = datetime.now()
    time = now.strftime("%H:%M:%S")
    return [fileNameDate, todayDate, time]

def printPaper(filePath, tempSelectionList, returnIndexValue):
    global loggedData
    if(systemPlatform):
        p.hw("INIT")
        printerStatus = p.paper_status()
        printerStatus = p.paper_status()
        if (printerStatus != 0):
            p.image(filePath)
    
            if((config_values["COBRANDING"] == "YES") and (int(config_values["RENEWAL_DATE"]) <= int(getCurrentDate()))):
                if(os.path.exists(os.path.join(ICONS_ROOT, "Cobranding_NoLogoAvailable.png"))):
                    p.image(os.path.join(ICONS_ROOT, "Cobranding_NoLogoAvailable.png"))

            elif((config_values["COBRANDING"] == "YES") and (int(config_values["RENEWAL_DATE"]) > int(getCurrentDate()))):
                if(os.path.exists(os.path.join(ICONS_ROOT, "cobranding.png"))):
                    p.image(os.path.join(ICONS_ROOT, "cobranding.png"))
                
            x = datetime.now()
            fullDateprint = str(x.strftime("%d/%m/%Y - %H:%M"))

            p.textln(fullDateprint)

            if("englishDash_storyPoem" in filePath):
                p.textln(tempSelectionList[2] + "/" +
                        tempSelectionList[3])
            """
            p.textln(fullDateprint + " - " +
                    tempSelectionList[0] + "/" +
                    tempSelectionList[1] + "/" +
                    tempSelectionList[2] + "/" +
                    tempSelectionList[3])
            """
            p.image(os.path.join(ICONS_ROOT, "footer.png"))
            p.cut()

            tempTime = getCurrent_DateTime()
            saveExcell_Data().saveExcell_DataCheck(tempTime, tempSelectionList, returnIndexValue)
        else:
            print("Error_No Paper")
            systemLogger.error("Error_No Paper")
            loggedData = "PaperJam"
            infoScreen("", [], "error_paperError", "")
    else:
        systemLogger.info("App running in WINDOWS.")
        systemLogger.info(tempSelectionList)

def checkReminderNotification(renewalDate):
    tempRenewalDate = datetime.strptime(renewalDate, "%Y%m%d")
    todayDate = datetime.now()

    delta = tempRenewalDate - todayDate
    if delta.days < reminderDays:
        tempReturn = True
    else:
        tempReturn = False
    return tempReturn

def checkRentalSubscription():
    global loggedData
    
    main_page.pack_forget()
    if(loggedData == "PaperJam"):
        print("PaperJam")
        window.after(5, infoScreen, "", [], "error_paperError", "")
    else:
        now = datetime.now()
        tempDate = now.strftime("%Y%m%d")

        if(str(config_values["RENTAL_OR_SUBSCRIPTION"]) == "BASIC_SUBSCRIPTION" or str(config_values["RENTAL_OR_SUBSCRIPTION"]) == "FULL_SUBSCRIPTION"):
            if(int(config_values["RENEWAL_DATE"]) < int(tempDate)):
                systemLogger.debug("subscription Renewal Due")
                window.after(5, infoScreen, "", [], "error_subscriptionRenewal", "")
            else:
                if(checkReminderNotification(str(int(config_values["RENEWAL_DATE"])))):
                    systemLogger.debug('subscriptionRenewal')
                    window.after(5, infoScreen, "", [], "error_subscriptionShortly", "")
                else:
                    print("enterSUB Not Due")
                    window.after(5, iconDashboard, main_page, LANGUAGE_LIST, "first Page")

        elif(str(config_values["RENTAL_OR_SUBSCRIPTION"]) == "RENTAL"):
            if(int(config_values["RENEWAL_DATE"]) < int(tempDate)):
                systemLogger.debug("SYSTEM LOCKED")
                window.after(5, infoScreen, "", [], "error_rentalRenewal", "")
            else:
                if(checkReminderNotification(str(int(config_values["RENEWAL_DATE"])))):
                    systemLogger.debug('rentalRenewal')
                    window.after(5, infoScreen, "", [], "error_rentalRenewalShortly", "")
                else:
                    print("enterRental Not Due")
                    window.after(5, iconDashboard, main_page, LANGUAGE_LIST, "first Page")

#Back Button Code for ALL SCREENS
def iconBack(screenName):
    screenName = screenName.split()
    print(screenName[0])

    if("DASH_LIST" in screenName[0].upper()): 
        main_page.pack_forget()
        iconDashboard(main_page, LANGUAGE_LIST, "first Page")

    elif("EVENT" in screenName[0].upper()):
        main_page.pack_forget()
        iconDashboard(main_page, LANGUAGE_LIST, "first Page")
    
    elif("DASH_STORYPOEM" in screenName[0].upper()): 
        main_page.pack_forget()
        tempName = screenName[0].upper().split("_")
        listName = tempName[0].upper() + "_LIST"
        iconDashboard(main_page, eval(listName), listName)

    elif("GENRE_LIST" in screenName[0].upper()): 
        main_page.pack_forget()
        listName = screenName[0].upper().split("GENRE_LIST")
        iconDashboard(main_page, eval(listName[0] + "DASH_LIST") , listName[0] + "DASH_LIST" )
    
    elif(("GENRE_STORY" in screenName[0].upper()) or
            ("GENRE_POEM" in screenName[0].upper()) or
            ("GENRE_ANYTHING" in screenName[0].upper())): 
        main_page.pack_forget()
        tempName = screenName[0].upper().split("_")
        listName = tempName[0].upper() + "_LIST"
        iconDashboard(main_page, eval(listName), listName)

    elif("DASH_BOOKINPARTS" in screenName[0].upper()): 
        main_page.pack_forget()
        tempName = screenName[0].upper().split("_")
        listName = tempName[0].upper() + "_LIST"
        iconDashboard(main_page, eval(listName), listName)

    elif("DASH_INTERESTINGFACTS" in screenName[0].upper()): 
        main_page.pack_forget()
        tempName = screenName[0].upper().split("_")
        listName = tempName[0].upper() + "_LIST"
        iconDashboard(main_page, eval(listName), listName)

    elif("DASH_PUZZLES" in screenName[0].upper()): 
        main_page.pack_forget()
        tempName = screenName[0].upper().split("_")
        listName = tempName[0].upper() + "_LIST"
        iconDashboard(main_page, eval(listName), listName)

    elif("PUZZLES_LIST" in screenName[0].upper()): 
        main_page.pack_forget()
        listName = screenName[0].upper().split("PUZZLES_LIST")
        iconDashboard(main_page, eval(listName[0] + "DASH_LIST") , listName[0] + "DASH_LIST" )

    elif(("PUZZLES_NUMBERPUZZLES" in screenName[0].upper()) or
               ("PUZZLES_WORDSEARCH" in screenName[0].upper())or
               ("PUZZLES_CROSSWORD" in screenName[0].upper())or
               ("PUZZLES_KINDERGARTEN" in screenName[0].upper())):
        main_page.pack_forget()
        tempName = screenName[0].upper().split("_")
        listName = tempName[0].upper() + "_LIST"
        iconDashboard(main_page, eval(listName), listName)


#UI Button Events for ALL SCREENS
def iconDashboard_Clicked(buttonName):
    global LANGUAGE_SELECTION
    global SUB_ONE_SELECTION
    global SUB_TWO_SELECTION
    global SUB_THREE_SELECTION
    print(buttonName)

    if(buttonName in LANGUAGE_LIST):
        LANGUAGE_SELECTION = buttonName
        print("LANGUAGE_SELECTION :" + LANGUAGE_SELECTION)
        listName = LANGUAGE_SELECTION.upper()+"DASH_LIST"
        main_page.pack_forget()
        iconDashboard(main_page, eval(listName), listName)

    elif(buttonName in EVENT_LIST):
        SUB_ONE_SELECTION = buttonName
        print(LANGUAGE_SELECTION +"/"+ SUB_ONE_SELECTION+"/")
        printSettings(LANGUAGE_SELECTION,
                SUB_ONE_SELECTION,
                "",
                "")
        
    elif((buttonName in TELUGUDASH_LIST) or 
            (buttonName in HINDIDASH_LIST) or 
            (buttonName in ENGLISHDASH_LIST) ):
        SUB_ONE_SELECTION = buttonName
        print("SUB_ONE_SELECTION :" + LANGUAGE_SELECTION +"/"+ SUB_ONE_SELECTION)

        if("Dash_storyPoem" in SUB_ONE_SELECTION):
            listName = SUB_ONE_SELECTION.split("Dash_storyPoem")[0]+"GENRE_LIST"
            main_page.pack_forget()
            iconDashboard(main_page, eval(listName.upper()), SUB_ONE_SELECTION)

        elif("Dash_bookInParts" in SUB_ONE_SELECTION):
            listName = SUB_ONE_SELECTION.split("Dash_bookInParts")[0]+"BIP_LIST"
            main_page.pack_forget()
            iconDashboard(main_page, eval(listName.upper()), SUB_ONE_SELECTION)


        elif("Dash_interestingFacts" in SUB_ONE_SELECTION):
            listName = SUB_ONE_SELECTION.split("Dash_interestingFacts")[0]+"FACTS_LIST"
            main_page.pack_forget()
            iconDashboard(main_page, eval(listName.upper()), SUB_ONE_SELECTION)
        

        elif("Dash_puzzles" in SUB_ONE_SELECTION):
            listName = SUB_ONE_SELECTION.split("Dash_puzzles")[0]+"PUZZLES_LIST"
            main_page.pack_forget()
            iconDashboard(main_page, eval(listName.upper()), SUB_ONE_SELECTION)

    elif((buttonName in TELUGUGENRE_LIST) or 
            (buttonName in TELUGUPUZZLES_LIST) or 
            (buttonName in TELUGUBIP_LIST) or 
            (buttonName in TELUGUFACTS_LIST) or

            (buttonName in HINDIGENRE_LIST) or
            (buttonName in HINDIPUZZLES_LIST) or
            (buttonName in HINDIBIP_LIST) or
            (buttonName in HINDIFACTS_LIST) or

            (buttonName in ENGLISHGENRE_LIST) or
            (buttonName in ENGLISHPUZZLES_LIST) or
            (buttonName in ENGLISHBIP_LIST) or
            (buttonName in ENGLISHFACTS_LIST)):
        SUB_TWO_SELECTION = buttonName
        print("SUB_TWO_SELECTION :" + LANGUAGE_SELECTION +"/"+ SUB_ONE_SELECTION+"/"+SUB_TWO_SELECTION)

        if(("Genre_story" in SUB_TWO_SELECTION) or 
                ("Genre_poem" in SUB_TWO_SELECTION) or 
                ("Genre_anything" in SUB_TWO_SELECTION)):
            listName = SUB_TWO_SELECTION.split("Genre_")[0]+"AGE_LIST"
            main_page.pack_forget()
            iconDashboard(main_page, eval(listName.upper()), SUB_TWO_SELECTION)

        elif(("Puzzles_numberPuzzles" in SUB_TWO_SELECTION) or
                ("Puzzles_wordSearch" in SUB_TWO_SELECTION) or
                ("Puzzles_crossWord" in SUB_TWO_SELECTION) or
                ("Puzzles_kindergarten" in SUB_TWO_SELECTION)):
            listName = SUB_TWO_SELECTION.split("_")[0]+"AGE_LIST"
            main_page.pack_forget()
            iconDashboard(main_page, eval(listName.upper()), SUB_TWO_SELECTION)

        elif(("bookOne" in buttonName) or
                ("bookTwo" in buttonName) or
                ("bookThree" in buttonName) or ("bookFour" in buttonName)):
            print(LANGUAGE_SELECTION +"/"+ SUB_ONE_SELECTION+"/"+SUB_TWO_SELECTION)
            printSettings(LANGUAGE_SELECTION,
                  SUB_ONE_SELECTION,
                  SUB_TWO_SELECTION,
                  "")
        
        elif(("factsBiographies" in buttonName) or
                ("factsScience" in buttonName) or
                ("factsHistory" in buttonName)):
            print(LANGUAGE_SELECTION +"/"+ SUB_ONE_SELECTION+"/"+SUB_TWO_SELECTION)
            printSettings(LANGUAGE_SELECTION,
                  SUB_ONE_SELECTION,
                  SUB_TWO_SELECTION,
                  "")

    elif((buttonName in TELUGUAGE_LIST) or 
        (buttonName in TELUGUPUZZLESAGE_LIST) or

        (buttonName in HINDIAGE_LIST) or
        (buttonName in HINDIPUZZLESAGE_LIST) or

        (buttonName in ENGLISHAGE_LIST) or
        (buttonName in ENGLISHPUZZLESAGE_LIST)):
        SUB_THREE_SELECTION = buttonName
        print("SUB_THREE_SELECTION :" + LANGUAGE_SELECTION +"/"+ SUB_ONE_SELECTION+"/"+SUB_TWO_SELECTION+"/"+SUB_THREE_SELECTION)
        printSettings(LANGUAGE_SELECTION,
                  SUB_ONE_SELECTION,
                  SUB_TWO_SELECTION,
                  SUB_THREE_SELECTION)

######################################################
##******************* CHECK INTERNET ***************##

def ping_checkInternet():
    try:
        socket.setdefaulttimeout(3)
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        host = "8.8.8.8"
        port = 53
        server_address = (host, port)
        s.connect(server_address)
    except Exception as error:
        systemLogger.debug(f"No Internet Connected - {str(error)}")
        print('No Internet Connected')
        return False
    else:
        s.close()
        systemLogger.debug('Internet Connected')
        print('Internet Connected')
        return True

######################################################

def event_clicked(event):
    global LANGUAGE_SELECTION
    LANGUAGE_SELECTION = "CustomeEvent"
    iconDashboard(main_page, EVENT_LIST, "event page") 

def iconDashboard(page_name, temp_list, pageName):
    try:
        lightsColor("White")
        main_canvas = Canvas(page_name,
                            bg = whiteColor,
                            width = windowWidth,
                            height = windowHeight,
                            bd = 0,
                            highlightthickness = 0,
                            relief = "ridge")
        main_canvas.place(x = 0, y = 0)

        if(checkReminderNotification(str(int(config_values["RENEWAL_DATE"])))):
            img= (Image.open(os.path.join(ICONS_ROOT, "error_renewalDue.png")))
            storyboxRenewal_image= img.resize((160,80), Image.Resampling.LANCZOS)
            storyboxRenewal_image= ImageTk.PhotoImage(storyboxRenewal_image)
            main_canvas.create_image(120, 410,
                                        image=storyboxRenewal_image,
                                        anchor=CENTER)

        if(systemPlatform):
            if(paperSensor.value == False):
                img= (Image.open(os.path.join(ICONS_ROOT, "error_paperLow.png")))
                paperLow_image= img.resize((120,80), Image.Resampling.LANCZOS)
                paperLow_image= ImageTk.PhotoImage(paperLow_image)
                main_canvas.create_image(700, 70,
                                            image=paperLow_image,
                                            anchor=CENTER)
        
        if(pageName == "first Page"):
            if(os.path.isfile(os.path.join(PROJECT_ROOT,"qrCode.png"))):
                os.remove(os.path.join(PROJECT_ROOT,"qrCode.png"))

            today = int(datetime.now().strftime("%Y%m%d"))
            if((config_values["EVENT_TYPE"] == "YES") and (config_values["EVENT_NUMBER"] >= today)):
                event_img = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "event.png")))
                event = main_canvas.create_image(400, 90,
                                                    image=event_img)
                main_canvas.tag_bind(event, "<Button-1>", event_clicked)

        if(pageName != "first Page"):
            img= (Image.open(os.path.join(ICONS_ROOT, "back.png")))
            back_image= img.resize((backButtonWidthHeight,backButtonWidthHeight),
                                Image.Resampling.LANCZOS)
            back_image = ImageTk.PhotoImage(back_image)
            Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = back_image, 
                    command= lambda :iconBack(pageName+" back")).place(x = 30 ,
                                                               y = 30)

        quitButton_img = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, "error_quit.png")))
        Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = quitButton_img, 
                    command= lambda :quitButton()).place(x = 795 ,
                                                               y = 0)

        i=0
        source_img = []
        icon_list = []
        if(pageName == "event page"):
            source_img = [sub for sub in temp_list]
            x = 0
            for x in range (int(config_values["EVENT_NUMBER"])):
                icon_list.append(source_img[x])
            #print(icon_list)
        else:
            source_img = [sub + "_img" for sub in temp_list]
            icon_list = temp_list

        for i in range(len(icon_list)):
            source_img[i] = ImageTk.PhotoImage(Image.open(os.path.join(ICONS_ROOT, icon_list[i]+".png")))

            if(len(icon_list) == 1):
                Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = source_img[i], 
                    command=lambda i=i: iconDashboard_Clicked(icon_list[i]) ).place(x = buttonAlign_X +(300) ,
                                                                                                            y = buttonAlign_Y)

            if(len(icon_list) == 2):
                Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = source_img[i], 
                    command=lambda i=i: iconDashboard_Clicked(icon_list[i]) ).place(x = buttonAlign_X +(200)+(210*i) ,
                                                                                                            y = buttonAlign_Y)

            elif(len(icon_list) == 3):
                Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = source_img[i], 
                    command=lambda i=i: iconDashboard_Clicked(icon_list[i]) ).place(x = buttonAlign_X +(50)+(230*i) ,
                                                                                                            y = buttonAlign_Y)

            elif(len(icon_list) == 4):
                Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = source_img[i], 
                    command=lambda i=i: iconDashboard_Clicked(icon_list[i]) ).place(x = buttonAlign_X +(175*i)+(20*i) ,
                                                                                                            y = buttonAlign_Y)

            elif(len(icon_list) == 5):
                Button(main_canvas, bg=whiteColor, borderwidth=0, activebackground=whiteColor, highlightthickness = 0,
                    image = source_img[i], 
                    command=lambda i=i: iconDashboard_Clicked(icon_list[i]) ).place(x = buttonAlign_X + (135*i) + (20*i) ,
                                                                                                            y = buttonAlign_Y)

        if(pageName == "first Page"):
            wifi_image = ""
            if(ping_checkInternet()):
                img= Image.open(os.path.join(ICONS_ROOT, "error_wifi.png"))
            else:
                img= Image.open(os.path.join(ICONS_ROOT, "error_wifiNo.png"))

            wifi_image= img.resize((60,60), Image.Resampling.LANCZOS)
            nowifi_image= ImageTk.PhotoImage(wifi_image)
            main_canvas.create_image(700, 440,
                                        image=nowifi_image,
                                        anchor=CENTER)   

        storyboxLogo = ImageTk.PhotoImage(Image.open(ICONS_ROOT + "/error_storyboxLogo.png"))
        label1 = Label(main_canvas, image=storyboxLogo, bg='white', borderwidth=0)
        label1.image = storyboxLogo
        label1.place(x=400, y = 440, anchor=CENTER)

        raise_frame(page_name)

        #Turning ON HDMI Screen
        if(pageName == "first Page"):
            os.system('vcgencmd display_power 1')
        window.mainloop()
    except Exception as error:
        systemLogger.error(f"Unable to open MAIN SCREEN UI - {str(error)}")
        #page_name.pack_forget()
        infoScreen("", [], "error_screen", "")
        #window.after(10,errorScreenFunction)


if __name__ == "__main__":
    now = datetime.now()
    tempDate = now.strftime("%Y%m%d")

    #if (((p == "Error") or (device is None)) and (systemPlatform)):
    if (((p == "E") and (device is None)) and (systemPlatform)):
        infoScreen("", [], "error_printerCommunicationError", "")

    else:
        if((int(config_values["RENEWAL_DATE"]) < int(tempDate)) and (str(config_values["RENTAL_OR_SUBSCRIPTION"]) == "RENTAL")):
            print("LOCKED")
            infoScreen("", [], "error_rentalRenewal", "")
        else:
            iconDashboard(main_page, LANGUAGE_LIST, "first Page")
        
