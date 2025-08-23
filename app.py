import sys
import os
from openpyxl import load_workbook
from xlutils.copy import copy as xl_copy
from setup_gui import setup, get_setup_info
from excel_processor import *
from config_utils import *
from setup_gui import *
from helpers import *

APP_NAME = "Cutting Generator"
APP_VERSION = "v2.0.0"
AUTHOR = "Lian Jordaan"

WINDOW_TITLE = f"{APP_NAME} {APP_VERSION} - {AUTHOR}"

if getattr(sys, 'frozen', False):
    # Running as a PyInstaller EXE
    base_path = sys._MEIPASS
else:
    # Running as a normal script
    base_path = os.path.dirname(__file__)

# Set terminal window title (works on most terminals)

if os.name == "nt":
    os.system(f"title {WINDOW_TITLE}")
else:
    sys.stdout.write(f"\x1b]2;{WINDOW_TITLE}\x07")


if __name__ == "__main__":

    config = get_setup_info()
    if not config:
        setup()
        config = get_setup_info()


    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        template_path = os.path.join(base_path, "template.xls")
        
        customer_name = get_customer_name(file_path)
        job_name = get_job_name(file_path)

        process_excel(file_path, template_path)

        print(f"Processed file: {file_path}")
        print("Attempting to search for cutouts...")
        try:
            cutouts = find_cutouts(customer_name, job_name)
            if cutouts:
                print("✅ Cutouts found:")
                for cutout in cutouts:
                    print(cutout)
            else:
                print("❌ No cutouts found for this customer and job.")
        except Exception as e:
            print(f"❌ Error while searching for cutouts: {e}")
            print("Please ensure the database configuration is correct in the setup.")
        
        input("Done. You may now close this window...")
        refresh_desktop()
    else:
        print("No file provided. Please drag an Excel file onto this program.")
