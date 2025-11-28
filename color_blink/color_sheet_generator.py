import xlwings as xw
import time
import random
import argparse

# -------- DEFAULT CONFIG --------
DEFAULT_EXCEL_FILE = r"sample.xlsx"
DEFAULT_SHEET_NAME = "Sheet1"
DEFAULT_TARGET_CELL = "A1"
DEFAULT_INTERVAL_SEC = 1

DEFAULT_ROW_HEIGHT_POINTS = 12
DEFAULT_COLUMN_WIDTH_CHARS = 2

DEFAULT_ROWS_TO_SET = 50
DEFAULT_COLS_TO_SET = 64


def blink_excel_cell(
    excel_file=DEFAULT_EXCEL_FILE,
    sheet_name=DEFAULT_SHEET_NAME,
    target_cell=DEFAULT_TARGET_CELL,
    interval_sec=DEFAULT_INTERVAL_SEC,
    row_height=DEFAULT_ROW_HEIGHT_POINTS,
    col_width=DEFAULT_COLUMN_WIDTH_CHARS,
    rows=DEFAULT_ROWS_TO_SET,
    cols=DEFAULT_COLS_TO_SET,
):
    """
    Normal function call — no argparse.
    Opens Excel, makes selected cells square, and blinks target cell.
    """
    wb = xw.Book(excel_file)
    sht = wb.sheets[sheet_name]

    # Make cells square
    for r in range(1, rows + 1):
        sht.api.Rows(r).RowHeight = row_height
    for c in range(1, cols + 1):
        sht.api.Columns(c).ColumnWidth = col_width

    print(f"Formatted first {rows} rows × {cols} columns as squares.")
    print(f"Blinking cell {target_cell}. Press Ctrl+C to stop.")

    cell = sht.range(target_cell)
    try:
        while True:
            cell.color = (random.randint(0, 255),
                          random.randint(0, 255),
                          random.randint(0, 255))
            time.sleep(interval_sec)
    except KeyboardInterrupt:
        print("Stopped.")


def blink_excel_cell_argparse():
    """
    CLI execution version — accepts input through command line arguments.
    """
    parser = argparse.ArgumentParser(description="Blink a cell in Excel and format cells as square.")
    parser.add_argument("--excel_file", default=DEFAULT_EXCEL_FILE, help="Path to Excel file")
    parser.add_argument("--sheet_name", default=DEFAULT_SHEET_NAME, help="Sheet name")
    parser.add_argument("--target_cell", default=DEFAULT_TARGET_CELL, help="Cell to blink")
    parser.add_argument("--interval", type=float, default=DEFAULT_INTERVAL_SEC, help="Blink interval seconds")

    parser.add_argument("--row_height", type=float, default=DEFAULT_ROW_HEIGHT_POINTS, help="Row height in points")
    parser.add_argument("--col_width", type=float, default=DEFAULT_COLUMN_WIDTH_CHARS, help="Column width in chars")
    parser.add_argument("--rows", type=int, default=DEFAULT_ROWS_TO_SET, help="Number of rows to resize")
    parser.add_argument("--cols", type=int, default=DEFAULT_COLS_TO_SET, help="Number of columns to resize")

    args = parser.parse_args()

    blink_excel_cell(
        args.excel_file,
        args.sheet_name,
        args.target_cell,
        args.interval,
        args.row_height,
        args.col_width,
        args.rows,
        args.cols,
    )

