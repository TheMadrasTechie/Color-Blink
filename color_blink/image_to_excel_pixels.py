import xlwings as xw
from PIL import Image
import argparse
import os

# -------- DEFAULT CONFIG --------
DEFAULT_ROW_HEIGHT_POINTS = 9
DEFAULT_COLUMN_WIDTH_CHARS = 1


def image_to_excel_cell_by_cell(
    image_path,
    excel_file,
    sheet_name="Sheet1",
    row_height=DEFAULT_ROW_HEIGHT_POINTS,
    column_width=DEFAULT_COLUMN_WIDTH_CHARS
):
    """
    Convert image pixel colors into Excel cells (no argparse here).
    """
    if not os.path.exists(image_path):
        print(f"Image not found: {image_path}")
        return

    img = Image.open(image_path).convert("RGB")
    width, height = img.size
    print(f"Image size: {width} x {height} (W x H)")

    wb = xw.Book(excel_file) if os.path.exists(excel_file) else xw.Book()
    wb.save(excel_file)
    sht = wb.sheets[sheet_name]

    # apply custom or default cell sizes
    sht.range((1, 1), (height, 1)).row_height = row_height
    sht.range((1, 1), (1, width)).column_width = column_width

    print("Writing pixels to Excel (may take time)...")
    for y in range(height):
        for x in range(width):
            r, g, b = img.getpixel((x, y))
            sht.range((y + 1, x + 1)).color = (int(r), int(g), int(b))

    print(f"Completed → {excel_file}  (sheet: '{sheet_name}')")


def image_to_excel_argparse():
    """
    Same functionality but with command-line argument support.
    """
    parser = argparse.ArgumentParser(description="Convert image pixels to Excel cells")
    parser.add_argument("--image_path", required=True, help="Path to input image")
    parser.add_argument("--excel_file", required=True, help="Path to output Excel file")
    parser.add_argument("--sheet_name", default="Sheet1", help="Sheet name")
    parser.add_argument("--row_height", type=int, default=DEFAULT_ROW_HEIGHT_POINTS,
                        help=f"Row height in points (default: {DEFAULT_ROW_HEIGHT_POINTS})")
    parser.add_argument("--column_width", type=float, default=DEFAULT_COLUMN_WIDTH_CHARS,
                        help=f"Column width in Excel chars (default: {DEFAULT_COLUMN_WIDTH_CHARS})")

    args = parser.parse_args()

    image_to_excel_cell_by_cell(
        args.image_path,
        args.excel_file,
        args.sheet_name,
        args.row_height,
        args.column_width
    )