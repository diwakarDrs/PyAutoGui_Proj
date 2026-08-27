# Daily Report Bot

A Windows desktop automation script that:

1. Opens Chrome in guest mode and searches Google for the ITC share price.
2. Copies the price from the search result.
3. Opens Excel, types the data into a new workbook, formats it as a table, and saves it as `daily_report_<date>.xlsx`.
4. Takes a screenshot of the final result.

Built with [`pyautogui`](https://pyautogui.readthedocs.io/) for keyboard/mouse automation and [`pyperclip`](https://pypi.org/project/pyperclip/) for clipboard access.

## Requirements

- Windows (uses `win+r`, Windows-style Save As dialog, etc.)
- Google Chrome and Microsoft Excel installed
- Python 3.9+

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python daily_report_bot.py
```

Run it with Chrome and Excel closed beforehand for the most reliable behavior. The script will:

- Open Chrome and search "ITC share price" on Google.
- Try a list of hardcoded screen coordinates to find and copy the price shown in Google's finance widget.
- Open Excel, create a new workbook, and type the date/time, price, and a comment into `A1:C2`.
- Format the range as an Excel Table, bold the header row, and format the price cell as a number.
- Save the file via the Windows "Save As" dialog (`F12` → type full path → Enter).
- Screenshot the final sheet.

Output files land in `daily_report_output/`:

- `daily_report_<date>.xlsx`
- `daily_report_<date>_screenshot.png`
- Diagnostic screenshots (`google_itc_failed_screen.png`, `save_error_screen_v25.png`) if a step fails.

## Known limitations

- **Screen-coordinate dependent.** The Google price lookup and a few Excel formatting steps click/type at fixed screen coordinates or rely on specific menu shortcuts. These were tuned against one screen resolution and one Excel/locale version — they may need adjusting (`coordinates` list, keyboard shortcuts) on a different machine or display setup.
- **No error recovery mid-flow.** If Chrome, Excel, or a dialog doesn't appear when expected, the script will likely misfire on the next step rather than pause and retry.
- **Not headless.** This drives real UI windows; don't use the mouse/keyboard while it's running, and keep `pyautogui.FAILSAFE` in mind (moving the mouse to a screen corner aborts the script).

## Configuration

Key constants at the top of `daily_report_bot.py`:

| Variable | Purpose |
|---|---|
| `OUTPUT_DIR` | Folder where the `.xlsx` and screenshots are saved |
| `COMMENT` | Text written into the "Comment" column |
| `coordinates` | Fallback list of screen positions tried when locating the price on the Google results page |
