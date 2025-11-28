import xlwings as xw
import time
import random
import re

EXCEL_FILE = r"sample.xlsx"   # <-- put your Excel file path here
SHEET_NAME = "Sheet1"         # <-- change if needed
CELL_ADDR = "A1"              # <-- cell to color & make square
INTERVAL_SEC = 1              # seconds between color changes

# Approx visual "square" size (tweak as you like)
ROW_HEIGHT_POINTS = 40        # row height in points
COLUMN_WIDTH_CHARS = 8        # column width in "character" units


def make_cell_square(cell):
    """
    Make the given cell's row and column *approximately* square.
    Excel uses different units for row height (points) and column width
    (character-based), so we just pick values that look square.
    """
    cell.row_height = ROW_HEIGHT_POINTS
    cell.column_width = COLUMN_WIDTH_CHARS


def main():
    # Open workbook (or attach if already open)
    try:
        wb = xw.Book(EXCEL_FILE)
    except FileNotFoundError:
        print(f"File not found: {EXCEL_FILE}")
        return

    sht = wb.sheets[SHEET_NAME]
    cell = sht.range(CELL_ADDR)

    # Make that cell's row/column square once at start
    make_cell_square(cell)

    print(f"Changing color every {INTERVAL_SEC} second(s) in {CELL_ADDR}. Press Ctrl+C to stop.")
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
