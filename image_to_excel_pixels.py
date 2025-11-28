import xlwings as xw
from PIL import Image
import os

# -------- CONFIG --------
IMAGE_PATH = r"template_3_reduced.jpg"   # your image
EXCEL_FILE = r"pixel_art.xlsx"           # output Excel
SHEET_NAME = "Sheet1"

# Cell size (tweak until it looks square on your screen)
ROW_HEIGHT_POINTS = 9
COLUMN_WIDTH_CHARS = 1


def image_to_excel_cell_by_cell():
    # --- Load image ---
    if not os.path.exists(IMAGE_PATH):
        print(f"Image not found: {IMAGE_PATH}")
        return

    # Force RGB so every pixel is (R,G,B)
    img = Image.open(IMAGE_PATH).convert("RGB")
    width, height = img.size
    print(f"Image size: {width} x {height} (W x H), mode: {img.mode}")

    # --- Open/create workbook ---
    if os.path.exists(EXCEL_FILE):
        wb = xw.Book(EXCEL_FILE)
    else:
        wb = xw.Book()
        wb.save(EXCEL_FILE)

    sht = wb.sheets[SHEET_NAME]

    # --- Make cells square-ish for the needed area ---
    # Set row heights for rows 1..height
    sht.range((1, 1), (height, 1)).row_height = ROW_HEIGHT_POINTS
    # Set column widths for cols 1..width
    sht.range((1, 1), (1, width)).column_width = COLUMN_WIDTH_CHARS

    # --- Set color cell by cell (rock-solid, no shape issues) ---
    print("Writing pixels to Excel (this may take a bit if the image is large)...")
    for y in range(height):        # Excel row index = y+1
        for x in range(width):     # Excel col index = x+1
            r, g, b = img.getpixel((x, y))
            sht.range((y + 1, x + 1)).color = (int(r), int(g), int(b))

    print(f"Done. Wrote pixels into {EXCEL_FILE} in sheet '{SHEET_NAME}'")
    print(f"Used rows 1–{height}, columns 1–{width}.")


if __name__ == "__main__":
    image_to_excel_cell_by_cell()
