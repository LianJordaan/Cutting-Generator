import sys
import warnings
import os
import shutil
import xlrd
from openpyxl import load_workbook
from xlutils.copy import copy as xl_copy
import xlwt
import ctypes

APP_NAME = "Cutting Generator"
APP_VERSION = "v1.2"
AUTHOR = "Lian Jordaan"

valid_bord_types = {"plain boards", "grain boards"}
invalid_bord_names = {"own peen", "own grain", "top 600", "top 900"}


WINDOW_TITLE = f"{APP_NAME} {APP_VERSION} - {AUTHOR}"

# Constants for SHChangeNotify
SHCNE_ASSOCCHANGED = 0x8000000  # Notify file type associations have changed
SHCNF_FLUSH = 0x1000            # Perform a flush after the notification
SHCNF_IDLIST = 0x0              # We are sending a notification, not a folder


if getattr(sys, 'frozen', False):
    # Running as a PyInstaller EXE
    base_path = sys._MEIPASS
else:
    # Running as a normal script
    base_path = os.path.dirname(__file__)


# Suppress the openpyxl default style warning
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")

input_file_path = ""

# Set terminal window title (works on most terminals)

if os.name == "nt":
    os.system(f"title {WINDOW_TITLE}")
else:
    sys.stdout.write(f"\x1b]2;{WINDOW_TITLE}\x07")

def process_excel(file_path, template_path):
    print(f"[INFO] Starting processing of file: {file_path}")

    # Extract job name and customer
    print("[INFO] Reading job information from input file...")
    customer = get_cell_value(file_path, 1, 2)
    job_name = get_cell_value(file_path, 2, 2)
    print(f"[INFO] Customer: {customer}, Job Name: {job_name}")

    output_filename = f"{job_name}.xls"
    output_path = os.path.join(os.path.dirname(file_path), output_filename)

    # Copy template
    print(f"[INFO] Copying template '{template_path}' to '{output_path}'...")
    shutil.copyfile(template_path, output_path)
    print("[INFO] Template copy complete.")

    # Count valid sheets
    num_valid_sheets = count_valid_sheets(input_file_path)
    print(f"[INFO] Number of valid sheets detected: {num_valid_sheets}")

    # Open template for editing
    print("[INFO] Opening template for modification...")
    rb = xlrd.open_workbook(output_path, formatting_info=True)
    wb = xl_copy(rb)
    print("[INFO] Template loaded and ready for editing.")

    # Prepare print sheet
    wb_print = xlwt.Workbook()
    ws_print = wb_print.add_sheet("Print", True)
    print("[INFO] Print sheet created.")

    current_row_print = 0
    to_write = []

    style_border_all_thin = make_border_style(1, 1, 1, 1)
    style_border_all_thin_bold = make_border_style(1, 1, 1, 1, bold=True)

    # Write initial job info to print sheet
    print("[INFO] Adding job header to print sheet...")
    to_write.append((current_row_print, 0, "Job Name", style_border_all_thin))
    to_write.append((current_row_print, 2, job_name, style_border_all_thin_bold))
    current_row_print += 2

    reading_sheet_index = 0
    processed_sheets = 0

    # Process each valid sheet
    for sheet_index in range(min(get_sheet_count(input_file_path), rb.nsheets)):
        print(f"[INFO] Processing sheet {sheet_index + 1}/{num_valid_sheets}")
        ws = wb.get_sheet(processed_sheets)
        if not is_sheet_valid(input_file_path, sheet_index):
            print(f"[WARNING] Sheet {sheet_index + 1} is not valid, skipping...")
            continue

        # Write job info to template
        print(f"[INFO] Writing job name and customer to template sheet {sheet_index + 1}...")
        ws.write(1, 2, job_name, style_border_all_thin)
        ws.write(0, 2, customer, style_border_all_thin)

        # Read bord and edging info
        bord_type = get_cell_value(file_path, 3, 2, sheet_index)
        bord_name = get_cell_value(file_path, 3, 4, sheet_index)
        edging_type = get_cell_value(file_path, 4, 2, sheet_index)
        edging_color = get_cell_value(file_path, 4, 4, sheet_index)
        print(f"[INFO] Bord Type: {bord_type}, Bord Name: {bord_name}")
        print(f"[INFO] Edging Type: {edging_type}, Edging Color: {edging_color}")

        ws.write(2, 2, bord_type, style_border_all_thin)
        ws.write(2, 4, bord_name, style_border_all_thin)
        ws.write(3, 2, edging_type, make_border_style(1, 1, 5, 0))
        ws.write(3, 4, edging_color, make_border_style(1, 0, 5, 1))

        # Fill table data
        last_empty_row = get_last_nonempty_row(input_file_path, 1, 7, sheet_index)
        print(f"[INFO] Last non-empty row in sheet {sheet_index + 1}: {last_empty_row}")

        for loop_col in range(1, 8):
            for loop_row in range(7, last_empty_row + 1):
                value = get_cell_value(file_path, loop_row, loop_col, sheet_index)
                ws.write(loop_row-1, loop_col, value, style_border_all_thin)
        print(f"[INFO] Table data written to template sheet {sheet_index + 1}")

        # Process edging information
        unique_edging = []
        for loop_row in range(7, last_empty_row + 1):
            loop_edging_category = str(get_cell_value(file_path, loop_row, 8, sheet_index)).lower()
            loop_edging_name = str(get_cell_value(file_path, loop_row, 9, sheet_index)).lower()

            # Normalization
            loop_edging_category = loop_edging_category.replace("pvc", "")
            loop_edging_category = loop_edging_category.replace("0.4mm", "pvc")
            loop_edging_category = loop_edging_category.replace("3mm", "2x36")
            edging_string = f"{loop_edging_category} {loop_edging_name}".upper()

            if "NO EDGING" in edging_string:
                edging_string = "NO EDGING"

            ws.write(loop_row-1, 10, edging_string, style_border_all_thin)

            if edging_string not in unique_edging:
                unique_edging.append(edging_string)
        print(f"[INFO] Unique edgings for sheet {sheet_index + 1}: {unique_edging}")

        # Add cutlist headers to print sheet
        to_write.append((current_row_print, 0, f"Cutlist {sheet_index + 1}", style_border_all_thin))
        to_write.append((current_row_print, 1, "Bord", style_border_all_thin))
        to_write.append((current_row_print, 2, bord_name, style_border_all_thin_bold))
        to_write.append((current_row_print, 3, "Hoeveelheid", style_border_all_thin))
        current_row_print += 3

        # Add edging info to print sheet
        to_write.append((current_row_print, 0, "Edging", style_border_all_thin))
        for edging in unique_edging:
            to_write.append((current_row_print, 2, edging, style_border_all_thin_bold))
            current_row_print += 1
        current_row_print += 1
        
        processed_sheets += 1

    to_write.append((current_row_print, 0, "CUTOUTS", style_border_all_thin))

    # Convert list to dict for easy writing
    to_write_dict = {}
    for row, col, text, style in to_write:
        to_write_dict[(row, col)] = (text, style)

    # Write to print sheet with borders
    print("[INFO] Writing print sheet data with borders...")
    for i in range(5):
        for j in range(current_row_print + 1):
            if (j, i) in to_write_dict:
                text, style = to_write_dict[(j, i)]
                ws_print.write(j, i, text, style)
            else:
                ws_print.write(j, i, "", style_border_all_thin)

    # Set column widths
    print("[INFO] Setting column widths for print sheet...")
    set_column_width_px(ws_print, 0, 70)
    set_column_width_px(ws_print, 2, 200)
    set_column_width_px(ws_print, 3, 88)

    # Save template sheet
    print(f"[INFO] Saving modified template as {output_path}...")
    wb.save(output_path)

    # Save print sheet
    output_path_print = os.path.join(os.path.dirname(file_path), "PRINT " + output_filename.replace('.xls', '_print.xls'))
    print(f"[INFO] Saving print sheet as {output_path_print}...")
    wb_print.save(output_path_print)
    print("[INFO] Processing complete.")

    if os.name == "nt":
        import ctypes
        import time
        user32 = ctypes.windll.user32
        progman = user32.FindWindowW("Progman", None)
        if progman:
            # 0x111 = WM_COMMAND, 0x7402 = Refresh
            user32.SendMessageW(progman, 0x111, 0x7402, 0)


def set_column_width_px(ws, col_idx, pixels):
    ws.col(col_idx).width = int((pixels) / 7 * 256)

def get_cell_value(file_path, row, col, sheet_index=0):
    """
    Returns the value at (row, col) from the specified sheet of the input file.
    Row and col are zero-based. Sheet index is zero-based.
    """

    if file_path.lower().endswith('.xlsx'):
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.worksheets[sheet_index]
        # openpyxl uses 1-based indexing for cell access
        cell_value = ws.cell(row=row + 1, column=col + 1).value  # pyright: ignore[reportOptionalMemberAccess]
        return cell_value
    else:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(sheet_index)
        return sheet.cell_value(row, col)

def get_sheet_count(file_path):
    """
    Returns the number of sheets in the input file.
    """
    if file_path.lower().endswith('.xlsx'):
        wb = load_workbook(file_path, read_only=True, data_only=True)
        return len(wb.worksheets)
    else:
        book = xlrd.open_workbook(file_path)
        return book.nsheets

def is_sheet_valid(file_path, sheet_index=0):
    """
    Returns True if the sheet at sheet_index in the input file is valid.
    A valid sheet has 'Plain Boards' or 'Grain Boards' in cell C4 and a non-empty name in E4.
    """
    valid_types = valid_bord_types
    invalid_names = invalid_bord_names
    if file_path.lower().endswith('.xlsx'):
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.worksheets[sheet_index]
        value = ws['C4'].value
        if isinstance(value, str) and value.strip().lower() in valid_types:
            name_value = ws['E4'].value
            if isinstance(name_value, str) and name_value.strip().lower() not in invalid_names:
                return True
    else:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(sheet_index)
        value = sheet.cell_value(2, 2)
        if isinstance(value, str) and value.strip().lower() in valid_types:
            name_value = sheet.cell_value(2, 4)
            if isinstance(name_value, str) and name_value.strip().lower() not in invalid_names:
                return True
    return False

def count_valid_sheets(file_path):
    """
    Returns the number of sheets where cell C3 is 'Plain Boards' or 'Grain Boards'.
    """
    valid_types = valid_bord_types
    invalid_names = invalid_bord_names
    count = 0

    if file_path.lower().endswith('.xlsx'):
        wb = load_workbook(file_path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            value = ws['C4'].value
            if isinstance(value, str) and value.strip().lower() in valid_types:
                name_value = ws['E4'].value
                if isinstance(name_value, str) and name_value.strip().lower() not in invalid_names:
                    count += 1
    else:
        book = xlrd.open_workbook(file_path)
        for sheet in book.sheets():
            value = sheet.cell_value(2, 2)
            if isinstance(value, str) and value.strip().lower() in valid_types:
                name_value = sheet.cell_value(2, 4)
                if isinstance(name_value, str) and name_value.strip().lower() not in invalid_names:
                    count += 1
    return count

def get_last_nonempty_row(file_path, col, start_row, sheet_index=0):
    """
    Returns the row index (zero-based) of the last non-empty cell in the given column,
    starting from start_row, in the specified sheet.
    """
    if file_path.lower().endswith('.xlsx'):
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.worksheets[sheet_index]
        max_row = ws.max_row
        last_row = None
        for r in range(start_row + 1, max_row + 1):  # openpyxl is 1-based
            value = ws.cell(row=r, column=col + 1).value
            if value not in (None, ""):
                last_row = r - 1  # convert to zero-based
        return last_row
    else:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(sheet_index)
        nrows = sheet.nrows
        last_row = None
        for r in range(start_row, nrows):
            value = sheet.cell_value(r, col)
            if value not in ("", None):
                last_row = r
        return last_row

def make_border_style(top, right, bottom, left, font_name='Calibri', font_size=12, bold=False):
    """
    Returns an xlwt.XFStyle with the specified border styles, Calibri font, and font size 11.
    Border values: 0=No border, 1=Thin, 2=Medium, 3=Dashed, 4=Dotted, 5=Thick, etc.
    """
    style = xlwt.XFStyle()
    style.borders.top = top
    style.borders.right = right
    style.borders.bottom = bottom
    style.borders.left = left
    font = xlwt.Font()
    font.name = font_name
    font.height = font_size * 20  # xlwt uses twips (1/20 of a point)
    font.bold = bold
    style.font = font
    return style

def refresh_desktop():
    # SHChangeNotify parameters: SHCNE_ASSOCCHANGED, SHCNF_FLUSH, None, None
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_FLUSH, None, None)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        template_path = os.path.join(base_path, "template.xls")
        input_file_path = file_path  # Store the input file path for later use
        process_excel(file_path, template_path)
        refresh_desktop()
    else:
        print("No file provided. Please drag an Excel file onto this program.")
