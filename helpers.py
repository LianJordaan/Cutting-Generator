import ctypes
import os

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