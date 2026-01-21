import sys
import psutil
import threading
import keyboard
import os
import time
import pyautogui
import subprocess

from enum import Enum

try:
    import win32gui
except ImportError:
    win32gui = None

#########################################################################################################################################################################################
# Constants
######A###################################################################################################################################################################################

SYMBOLS_TO_DOWNLOAD = [
    "ASBP",
    "AVR",
    "BAOS",
    "BOXL",
    "BTTC",
    "CDT",
    "CRVS",
    "DRCT",
    "ERNA",
    "FRSX",
    "GLXG",
    "GNLN",
    "HUDI",
    "IBG",
    "ICON",
    "IMTE",
    "IVF",
    "KUST",
    "LGHL",
    "LVRO",
    "PAPL",
    "POLA",
    "RAPT",
    "SDST",
    "SEGG",
    "SHPH",
    "SLGB",
    "SONM",
    "TRON",
    "TWG"
]

SYMBOL_LOAD_TIME = 10
SYMBOL_SEARCH_WINDOW_LOAD_WAIT_TIME = 5
SYMBOL_SEARCH_DUMMY_ENTRY = 'A'
DELAY_DEFAULT = 0.5
TYPEWRITE_INTERVAL = 0.05
POST_RUN_SCRIPT = "../ChartFileRenamer/Main.py"

POSITION_SYMBOL_SEARCH_FILTER_STOCKS = (954, 538)
POSITION_FIRST_STOCK_IN_SYMBOL_SEARCH = (928, 624)
POSITION_SEARCH_BAR = (938, 486)

POSITION_MAIN_MENU_BUTTON = (2180,57)
POSITION_DOWNLOAD_CHART_DATA_MENU_OPTION = (2270, 266)

POSITION_TIME_FORMAT_SELECTOR = (1280, 813)
POSITION_TIME_FORMAT_ISO = (1280, 850)

POSITION_EXPORT_BUTTON = (1420, 870)

POSITION_CHART_TOP_LEFT = (618, 321)
POSITION_CHART_BOTTOM_LEFT = (654, 998)
POSITION_CHART_TOP_RIGHT = (1923, 324)
POSITION_CHART_BOTTOM_RIGHT = (1927, 957)

#########################################################################################################################################################################################
# Enums
#########################################################################################################################################################################################

class LayoutChartEnum(Enum):
    TOP_LEFT_CHART = POSITION_CHART_TOP_LEFT
    BOTTOM_LEFT_CHART = POSITION_CHART_BOTTOM_LEFT
    TOP_RIGHT_CHART = POSITION_CHART_TOP_RIGHT
    BOTTOM_RIGHT_CHART = POSITION_CHART_BOTTOM_RIGHT

    def ClickAtPosition(self):
        return self.value

#########################################################################################################################################################################################
# Functions
#########################################################################################################################################################################################

def IsTradingViewRunning():
    isProcessFound = False
    isWindowFound = False

    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and "tradingview" in proc.info["name"].lower():
                isProcessFound = True
                break
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if win32gui:
        def _enum(hwnd, _):
            nonlocal isWindowFound
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd) or ""
                if "tradingview" in title.lower():
                    isWindowFound = True

        win32gui.EnumWindows(_enum, None)

    return isProcessFound and isWindowFound

def SelectChartToDownload(layoutChart: LayoutChartEnum):
    pyautogui.click(*layoutChart.ClickAtPosition())
    time.sleep(DELAY_DEFAULT)

def OpenMainMenu():
    pyautogui.click(*POSITION_MAIN_MENU_BUTTON)
    time.sleep(DELAY_DEFAULT)

def OpenExportChartDataWindow():
    pyautogui.click(*POSITION_DOWNLOAD_CHART_DATA_MENU_OPTION)
    time.sleep(DELAY_DEFAULT)

def OpenTimeFormatSelector():
    pyautogui.click(*POSITION_TIME_FORMAT_SELECTOR)
    time.sleep(DELAY_DEFAULT)

def SelectTimeFormatISO():
    pyautogui.click(*POSITION_TIME_FORMAT_ISO)
    time.sleep(DELAY_DEFAULT)

def ClickExportButton():
    pyautogui.click(*POSITION_EXPORT_BUTTON)
    time.sleep(DELAY_DEFAULT)

def ClickStocksFilter():
    pyautogui.click(*POSITION_SYMBOL_SEARCH_FILTER_STOCKS)
    time.sleep(DELAY_DEFAULT)

def ClickSearchBar():
    pyautogui.click(*POSITION_SEARCH_BAR)
    time.sleep(DELAY_DEFAULT)

def OpenSymbolSearchWindow(windowLoadWaitTime=SYMBOL_SEARCH_WINDOW_LOAD_WAIT_TIME):
    pyautogui.typewrite(SYMBOL_SEARCH_DUMMY_ENTRY, interval=TYPEWRITE_INTERVAL)
    time.sleep(windowLoadWaitTime)
    ClickStocksFilter()
    ClickSearchBar()
    pyautogui.press('backspace')

def TypeSymbol(symbol: str):
    pyautogui.typewrite(symbol, interval=TYPEWRITE_INTERVAL)
    time.sleep(1)

def SelectFirstSymbol():
    pyautogui.click(*POSITION_FIRST_STOCK_IN_SYMBOL_SEARCH)
    time.sleep(SYMBOL_LOAD_TIME)

def LoadSymbol(symbol: str):
    SelectChartToDownload(LayoutChartEnum.TOP_LEFT_CHART)
    OpenSymbolSearchWindow()
    TypeSymbol(symbol)
    SelectFirstSymbol()

def DownloadChartCSV(layoutChart: LayoutChartEnum):
    SelectChartToDownload(layoutChart)
    OpenMainMenu()
    OpenExportChartDataWindow()
    OpenTimeFormatSelector()
    SelectTimeFormatISO()
    ClickExportButton()

def DownloadCharts(symbol: str):
    LoadSymbol(symbol)
    DownloadChartCSV(LayoutChartEnum.TOP_LEFT_CHART)
    DownloadChartCSV(LayoutChartEnum.BOTTOM_LEFT_CHART)
    DownloadChartCSV(LayoutChartEnum.TOP_RIGHT_CHART)
    DownloadChartCSV(LayoutChartEnum.BOTTOM_RIGHT_CHART)

def RunScript(script_path: str):
    subprocess.run(["python", script_path])

def ListenForEscape():
    keyboard.wait('esc')
    print("\n🚫 ESC pressed. Exiting now.")
    os._exit(0)

#########################################################################################################################################################################################
# Main
#########################################################################################################################################################################################

if __name__ == "__main__":
    threading.Thread(target=ListenForEscape, daemon=True).start()

    # if not IsTradingViewRunning():
    #     print("❌ TradingView is not running or its window not found. Exiting.")
    #     sys.exit(1)

    for symbol in SYMBOLS_TO_DOWNLOAD:
        DownloadCharts(symbol)

    RunScript(POST_RUN_SCRIPT)
