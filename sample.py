import xlwings as xw
import time
import random

EXCEL_FILE = r"sample.xlsx"   # <-- put your Excel file path here
SHEET_NAME = "Sheet1"         # <-- change if needed
CELL_ADDR = "A1"              # <-- cell to color
INTERVAL_SEC = 1              # seconds between color changes


def main():
    # Open workbook (or attach if already open)
    try:
        wb = xw.Book(EXCEL_FILE)
    except FileNotFoundError:
        print(f"File not found: {EXCEL_FILE}")
        return

    sht = wb.sheets[SHEET_NAME]
    cell = sht.range(CELL_ADDR)

    print("Changing color every second. Press Ctrl+C to stop.")
    try:
        while True:
            r = random.randint(0, 255)
            g = random.randint(0, 255)
            b = random.randint(0, 255)

            # xlwings uses (R, G, B) tuple for cell color
            cell.color = (r, g, b)

            time.sleep(INTERVAL_SEC)
    except KeyboardInterrupt:
        print("\nStopped by user.")


if __name__ == "__main__":
    main()
