"""
daily_report_bot.py

"""

import os
import re
import subprocess
import time
from datetime import datetime

import pyautogui
import pyperclip
import pygetwindow as gw
from openpyxl import Workbook, load_workbook


# =====================================================================
# STEP 0 - Settings
# =====================================================================
pyautogui.FAILSAFE = True   # move mouse to top-left corner = emergency stop
pyautogui.PAUSE = 0.4       # small pause after every PyAutoGUI action

CITY = "Bengaluru"
# format=%l:+%t+%C  ->  "Bengaluru: +27°C Partly cloudy"
URL = f"https://wttr.in/{CITY}?format=%l:+%t+%C"

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
TODAY = datetime.now().strftime("%Y-%m-%d")                 # YYYY-MM-DD
FILE_STEM = f"daily_report_{TODAY}"
EXCEL_PATH = os.path.join(OUTPUT_DIR, f"{FILE_STEM}.xlsx")
SCREENSHOT_PATH = os.path.join(OUTPUT_DIR, f"{FILE_STEM}.png")


# =====================================================================
# Helper functions
# =====================================================================
def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def wait_for_window(title_part, timeout=30):
    """
    Wait until a window whose title contains `title_part` exists,
    bring it to the front, maximize it and return it.
    This solves 'is the app open yet?' and 'is the right window focused?'.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        # Prefer the window that is already in front
        active = gw.getActiveWindow()
        if active and title_part.lower() in active.title.lower():
            win = active
        else:
            matches = [w for w in gw.getWindowsWithTitle(title_part) if w.title]
            win = matches[0] if matches else None

        if win:
            try:
                if win.isMinimized:
                    win.restore()
                win.activate()
            except Exception:
                pass  # activate() sometimes raises on Windows even when it worked
            try:
                win.maximize()
            except Exception:
                pass
            time.sleep(1)
            return win

        time.sleep(0.5)

    raise RuntimeError(f"Window containing '{title_part}' did not appear in {timeout}s")


def paste_text(text):
    """
    Put text on the clipboard and press Ctrl+V.
    pyautogui.write() can't type special characters like '°',
    so pasting is the reliable way to enter any text.
    """
    pyperclip.copy(text)
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "v")


# =====================================================================
# STEP 1 - Open Chrome and go to the website
# =====================================================================
def open_website():
    log("Opening Chrome...")
    subprocess.Popen("start chrome --new-window", shell=True)
    wait_for_window("Chrome", timeout=30)

    log(f"Going to {URL}")
    pyautogui.hotkey("ctrl", "l")   # jump to the address bar
    paste_text(URL)
    pyautogui.press("enter")
    time.sleep(4)                   # give the page time to load


# =====================================================================
# STEP 2 - Copy the important information from the page
# =====================================================================
def copy_weather(retries=3):
    for attempt in range(1, retries + 1):
        log(f"Copying weather text (attempt {attempt})...")
        pyperclip.copy("")              # clear clipboard so we know the copy worked
        pyautogui.hotkey("ctrl", "a")   # select all text on the page
        pyautogui.hotkey("ctrl", "c")   # copy it
        time.sleep(0.5)

        text = pyperclip.paste().strip()
        if "°" in text:                 # looks like a real temperature
            log(f"Got: {text}")
            return text

        log("Page not ready yet, waiting a bit longer...")
        time.sleep(3)                   # page loaded slowly -> wait and retry

    raise RuntimeError("Could not copy weather data from the page")


# =====================================================================
# STEP 3 - Write a short comment based on the data
# =====================================================================
def make_comment(weather_text):
    match = re.search(r"([+-]?\d+)\s*°", weather_text)
    if not match:
        return "Check the weather manually"

    temp = int(match.group(1))
    if temp >= 32:
        return "Hot day - stay hydrated"
    if temp >= 20:
        return "Good for outdoor activities"
    if temp >= 12:
        return "Pleasant but carry a light jacket"
    return "Cold - dress warmly"


# =====================================================================
# STEP 4 - Make sure today's Excel file exists, find the next empty row
# (openpyxl only creates the empty file with headers; the bot types the row)
# =====================================================================
def prepare_workbook():
    if not os.path.exists(EXCEL_PATH):
        log(f"Creating {os.path.basename(EXCEL_PATH)}")
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        ws.append(["Date & Time", "Weather Data", "Comment"])
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 35
        wb.save(EXCEL_PATH)

    wb = load_workbook(EXCEL_PATH)
    next_row = wb.active.max_row + 1
    wb.close()
    return next_row


# =====================================================================
# STEP 5 - Open the file in Excel
# =====================================================================
def open_excel():
    log("Opening Excel...")
    os.startfile(EXCEL_PATH)        # opens the file with Excel
    win = wait_for_window(FILE_STEM, timeout=45)
    time.sleep(2)                   # let Excel finish drawing the sheet
    return win


# =====================================================================
# STEP 6 - Type the new row
# =====================================================================
def type_row(row, weather, comment):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")   # generated at run time
    log(f"Writing row {row}: {now} | {weather} | {comment}")

    pyautogui.press("esc")           # make sure we're not editing a cell
    pyautogui.hotkey("ctrl", "g")    # Excel 'Go To' dialog
    time.sleep(1)
    pyautogui.write(f"A{row}")
    pyautogui.press("enter")
    time.sleep(0.5)

    paste_text(now)
    pyautogui.press("tab")           # move to column B
    paste_text(weather)
    pyautogui.press("tab")           # move to column C
    paste_text(comment)
    pyautogui.press("enter")


# =====================================================================
# STEP 7 - Save the file
# =====================================================================
def save_excel():
    log("Saving...")
    pyautogui.hotkey("ctrl", "s")
    time.sleep(2)


# =====================================================================
# STEP 8 - Screenshot of the final sheet
# =====================================================================
def take_screenshot():
    pyautogui.hotkey("ctrl", "home")  # scroll back to the top-left cell
    time.sleep(1)
    pyautogui.screenshot(SCREENSHOT_PATH)
    log(f"Screenshot saved: {SCREENSHOT_PATH}")


# =====================================================================
# Main - run every step in order
# =====================================================================
def main():
    log("Starting in 3 seconds - hands off the mouse and keyboard!")
    time.sleep(3)

    open_website()
    weather = copy_weather()
    comment = make_comment(weather)

    row = prepare_workbook()
    open_excel()
    type_row(row, weather, comment)
    save_excel()
    take_screenshot()

    log(f"Done! Excel file: {EXCEL_PATH}")


if __name__ == "__main__":
    try:
        main()
    except pyautogui.FailSafeException:
        log("Stopped: mouse moved to the corner (fail-safe).")
    except Exception as e:
        log(f"ERROR: {e}")