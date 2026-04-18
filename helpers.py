import ctypes
import os
import webbrowser

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
    # to check if it is eriks file, we need to check the if its am excell file, and if it is, we need to check that the cell at A1 contains "name" ignore case. It will be a xls file.
    if file_path.lower().endswith('.xls') or file_path.lower().endswith('.xlsx'):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True)
            ws = wb.active
            cell_value = ws['A1'].value
            if cell_value and isinstance(cell_value, str) and cell_value.strip().lower() == 'name':
                return True
        except Exception as e:
            print(f"Error checking if file is Erik's: {e}")
    return False

