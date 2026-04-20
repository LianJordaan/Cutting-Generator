import ctypes
import os
import webbrowser
import sys

# Constants for SHChangeNotify
SHCNE_ASSOCCHANGED = 0x8000000  # Notify file type associations have changed
SHCNF_FLUSH = 0x1000            # Perform a flush after the notification
SHCNF_IDLIST = 0x0              # We are sending a notification, not a folder

def refresh_desktop():
    # SHChangeNotify parameters: SHCNE_ASSOCCHANGED, SHCNF_FLUSH, None, None
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_FLUSH, None, None)

def find_quote_num(file_path):
    filename = os.path.basename(file_path)           # e.g., "123 - filename.xlsx"
    quote_num = filename.split(' ')[0]               # "123"
    return quote_num

def open_license_link():
    webbrowser.open("https://raw.githubusercontent.com/LianJordaan/Cutting-Generator/refs/heads/master/LICENSE.txt")

def is_erik_cutlist(file_path):
    # Erik files are old .xls files; check A1 for "name" using xlrd.
    if file_path.lower().endswith('.xls'):
        try:
            import xlrd
            book = xlrd.open_workbook(file_path)
            sheet = book.sheet_by_index(0)
            cell_value = sheet.cell_value(0, 0)
            if isinstance(cell_value, str) and 'name' in cell_value.strip().lower():
                return True
        except Exception as e:
            print(f"Error checking if file is Erik's: {e}")
    return False

