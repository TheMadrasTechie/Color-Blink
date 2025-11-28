import xlwings as xw
import time
import random

EXCEL_FILE = r"sample.xlsx"
SHEET_NAME = "Sheet1"
TARGET_CELL = "A1"
INTERVAL_SEC = 1

# Square dimensions
ROW_HEIGHT_POINTS = 12          # row height
COLUMN_WIDTH_CHARS = 2          # column width

ROWS_TO_SET = 50
COLS_TO_SET = 64


def make_selected_cells_square(sheet):
    # set first N rows
    for r in range(1, ROWS_TO_SET + 1):
        sheet.api.Rows(r).RowHeight = ROW_HEIGHT_POINTS

    # set first N columns
    for c in range(1, COLS_TO_SET + 1):
        sheet.api.Columns(c).ColumnWidth = COLUMN_WIDTH_CHARS


def main():
    wb = xw.Book(EXCEL_FILE)
    sht = wb.sheets[SHEET_NAME]

    make_selected_cells_square(sht)

    cell = sht.range(TARGET_CELL)
    print("Formatted first 50 rows × 60 columns as squares.")
    print("Now blinking A1. Press Ctrl+C to stop.")

    try:
        while True:
            cell.color = (
                random.randint(0, 255),
                random.randint(0, 255),
                random.randint(0, 255)
            )
            time.sleep(INTERVAL_SEC)
    except KeyboardInterrupt:
        print("Stopped.")


if __name__ == "__main__":
    main()
