import pyautogui
import pyperclip
import time
import os
import re
from datetime import datetime

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.8

# ============================================================
# SETTINGS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "daily_report_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

now = datetime.now()
today_date = now.strftime("%Y-%m-%d")
today_time = now.strftime("%Y-%m-%d %H:%M:%S")

filename = f"daily_report_{today_date}.xlsx"
excel_path = os.path.join(OUTPUT_DIR, filename)
screenshot_path = os.path.join(
    OUTPUT_DIR,
    f"daily_report_{today_date}_screenshot.png"
)

COMMENT = "ITC share price copied from Google."


# ============================================================
# HELPER
# ============================================================

def extract_price(text):
    patterns = [
        r"(?:₹|Rs\.?|INR)\s*([0-9][0-9,]*\.[0-9]{1,2})",
        r"([0-9][0-9,]*\.[0-9]{1,2})",
        r"(?:₹|Rs\.?|INR)\s*([0-9][0-9,]*)",
    ]

    for pattern in patterns:
        for value in re.findall(pattern, text, re.I):
            value = value.replace(",", "")
            try:
                number = float(value)
            except ValueError:
                continue

            if 100 <= number <= 1000:
                return value

    return None


# ============================================================
# 1. OPEN CHROME
# ============================================================

print("Opening Chrome...")

pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("chrome", interval=0.05)
pyautogui.press("enter")
time.sleep(5)

pyautogui.hotkey("alt", "space")
time.sleep(0.5)
pyautogui.press("x")
time.sleep(2)

print("Opening Chrome Guest Mode...")

pyautogui.click(1340, 50)
time.sleep(1)
pyautogui.click(1260, 155)
time.sleep(5)


# ============================================================
# 2. GOOGLE
# ============================================================

print("Opening Google...")

pyautogui.hotkey("ctrl", "l")
pyautogui.write("google.com", interval=0.05)
pyautogui.press("enter")
time.sleep(5)

print("Searching ITC share price...")

pyautogui.write("ITC share price", interval=0.05)
pyautogui.press("enter")
time.sleep(8)


# ============================================================
# 3. GET ITC PRICE
# ============================================================

print("Finding ITC price on Google...")

coordinates = [
    (300, 350),   # confirmed working on user's screen
    (235, 500),
    (235, 350),
    (350, 350),
    (300, 400),
    (350, 400),
    (300, 450),
    (350, 450),
    (450, 350),
    (550, 350),
    (650, 350),
    (900, 300),
    (1000, 300),
    (1100, 300),
]

itc_price = None

for attempt, (x, y) in enumerate(coordinates, 1):
    pyperclip.copy("")

    pyautogui.moveTo(x, y, duration=0.15)
    pyautogui.doubleClick(interval=0.12)
    time.sleep(0.4)

    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.7)

    copied = pyperclip.paste().strip()

    print(
        f"Attempt {attempt}: "
        f"({x}, {y}) -> {copied!r}"
    )

    candidate = extract_price(copied)

    if candidate:
        itc_price = candidate
        break

if not itc_price:
    diagnostic = os.path.join(
        OUTPUT_DIR,
        "google_itc_failed_screen.png"
    )
    pyautogui.screenshot(diagnostic)

    raise RuntimeError(
        "Could not copy ITC share price from Google.\n"
        f"Diagnostic screenshot: {diagnostic}"
    )

print(f"ITC share price captured: {itc_price}")


# ============================================================
# 4. MOVE MOUSE AWAY FROM CORNERS BEFORE EXCEL
# ============================================================

pyautogui.moveTo(700, 400, duration=0.2)


# ============================================================
# 5. OPEN EXCEL
# ============================================================

print("Opening Microsoft Excel...")

pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("excel", interval=0.05)
pyautogui.press("enter")

print("Waiting for Excel to start...")
time.sleep(10)


# ============================================================
# 6. CREATE BLANK WORKBOOK
# ============================================================

print("Creating a new blank workbook...")

pyautogui.hotkey("ctrl", "n")
time.sleep(6)

pyautogui.press("esc")
pyautogui.hotkey("ctrl", "home")
time.sleep(1)


# ============================================================
# 7. POPULATE EXCEL WITHOUT CLIPBOARD PASTE
# ============================================================

print("Populating Excel cells...")

# A1
pyautogui.write("Date & Time", interval=0.05)
pyautogui.press("tab")

# B1
pyautogui.write("ITC Stock Price", interval=0.05)
pyautogui.press("tab")

# C1
pyautogui.write("Comment", interval=0.05)

pyautogui.press("home")
pyautogui.press("down")

# A2
pyautogui.write(today_time, interval=0.03)
pyautogui.press("tab")

# B2
pyautogui.write(itc_price, interval=0.05)
pyautogui.press("tab")

# C2
pyautogui.write(COMMENT, interval=0.03)

time.sleep(2)

print("Excel data populated successfully.")


# ============================================================
# 8. FORMAT EXCEL
# ============================================================

print("Formatting Excel sheet...")

pyautogui.hotkey("ctrl", "home")
pyautogui.keyDown("shift")
pyautogui.press("down")
pyautogui.press("right", presses=2, interval=0.15)
pyautogui.keyUp("shift")

time.sleep(0.5)

pyautogui.hotkey("ctrl", "t")
time.sleep(2)
pyautogui.press("enter")
time.sleep(2)

pyautogui.hotkey("ctrl", "home")
pyautogui.hotkey("ctrl", "space")
pyautogui.keyDown("shift")
pyautogui.press("right", presses=2, interval=0.15)
pyautogui.keyUp("shift")

pyautogui.hotkey("alt", "h")
time.sleep(0.4)
pyautogui.press("o")
time.sleep(0.4)
pyautogui.press("i")
time.sleep(2)

pyautogui.hotkey("ctrl", "home")
pyautogui.keyDown("shift")
pyautogui.press("right", presses=2, interval=0.15)
pyautogui.keyUp("shift")
pyautogui.hotkey("ctrl", "b")
time.sleep(1)

pyautogui.hotkey("ctrl", "home")
pyautogui.press("down")
pyautogui.press("right")
pyautogui.hotkey("ctrl", "shift", "1")
time.sleep(1)

pyautogui.hotkey("ctrl", "home")

print("Excel formatting completed.")


# ============================================================
# 9. SAVE - KEYBOARD ONLY, NO COORDINATE CLICKS
# ============================================================
#
# The previous version clicked hardcoded pixel positions to
# open the File menu, click Browse, click the File name field,
# and click Save. Any mismatch in screen resolution, DPI
# scaling, or dialog load timing made those clicks miss, which
# is why the file was never created.
#
# F12 opens the classic Windows "Save As" dialog directly, and
# that dialog always starts with focus in the File name field.
# Typing the FULL absolute path (folder + filename) into that
# field and pressing Enter saves straight to that exact
# location - no folder navigation, no Browse button, no mouse
# clicks required at all.

print("Opening Save As dialog with F12...")

pyautogui.press("f12")
time.sleep(3)

# Focus should already be in the File name field, but make sure
# by explicitly selecting all text in whatever is focused.
pyautogui.hotkey("ctrl", "a")
time.sleep(0.3)

pyautogui.write(excel_path, interval=0.02)
time.sleep(1)

print(f"Full save path entered: {excel_path}")

pyautogui.press("enter")
time.sleep(4)

# If a "file already exists, overwrite?" prompt appears, confirm it.
pyautogui.press("enter")
time.sleep(2)

# If a "keep current format?" prompt appears (rare for a plain
# .xlsx save), confirm it too.
pyautogui.press("enter")
time.sleep(2)


# ============================================================
# VERIFY FILE
# ============================================================

print("Checking whether the Excel file was saved...")

saved = False

for check in range(1, 11):

    if os.path.isfile(excel_path):
        saved = True

        print(
            f"Excel file confirmed on check {check}:"
            f"\n{excel_path}"
        )

        break

    print(
        f"File not found yet - check {check}/10..."
    )

    time.sleep(2)


if not saved:

    save_error_screen = os.path.join(
        OUTPUT_DIR,
        "save_error_screen_v25.png"
    )

    pyautogui.screenshot(save_error_screen)

    raise RuntimeError(
        "SAVE FAILED: Excel file was not created.\n\n"
        f"Expected:\n{excel_path}\n\n"
        f"Diagnostic screenshot:\n{save_error_screen}"
    )

print("Excel file saved successfully.")


# ============================================================
# 12. SCREENSHOT
# ============================================================

print("Taking final Excel screenshot...")

pyautogui.hotkey("ctrl", "home")
time.sleep(1)

pyautogui.screenshot(screenshot_path)

print("")
print("========================================")
print(" DAILY REPORT BOT COMPLETED")
print("========================================")
print(f"ITC Share Price : {itc_price}")
print(f"Excel File      : {excel_path}")
print(f"Screenshot      : {screenshot_path}")
print("========================================")