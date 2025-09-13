print("🚀 Starting Cutting Generator...")

import sys
import os

print("Loading libraries...")

from openpyxl import load_workbook
from xlutils.copy import copy as xl_copy
from setup_gui import setup, get_setup_info
from excel_processor import *
from config_utils import *
from setup_gui import *
from helpers import *
from shape_gen import *

print("Libraries loaded.")

APP_NAME = "Cutting Generator"
APP_VERSION = "v3.0.0"
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
        print("Setup complete. Please re-run the program.")
        print("Press enter to exit...")
        input()
        sys.exit(0)
    
    if not config.get("agree_terms"):
        print("You must agree to the terms and conditions before using this software.")
        print("Please re-run the setup and agree to the terms.")
        print("Or type YES to automatically agree to the terms and conditions and continue...")

        while True:
            choice = input().strip().upper()
            if choice == "YES" or choice == "Y":
                config["agree_terms"] = True
                save_config(config)
                print("Thank you. Your agreement has been saved. Please re-run the program.")
                break
            else:
                print("Please type YES to agree to the terms and conditions, or close the program.")
        sys.exit(0)



    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        template_path = os.path.join(base_path, "template.xls")
        
        qoute_num = find_quote_num(file_path)
        job_name = get_job_name(file_path)

        process_excel(file_path, template_path)

        print(f"Processed file: {file_path}")
        print("Attempting to search for cutouts...")
        try:
            cutouts = find_cutouts(qoute_num)
            if cutouts:
                print("✅ Cutouts found:")
                parsed_cutouts = []

                for cutout in cutouts:
                    length = cutout[2]
                    width = cutout[3]
                    amount = cutout[4]
                    code = cutout[5]

                    shape_id = code[:2]
                    value1 = int(code[2:6])
                    value2 = int(code[6:]) 

                    parsed_cutouts.append((shape_id, length, width, amount, value1, value2))
                
                output_filename = f"SHAPES {job_name}.pdf"
                output_path = os.path.join(os.path.dirname(file_path), output_filename)

                shapes_to_pdf(parsed_cutouts, output_pdf=output_path)

            else:
                print("❌ No cutouts found for this customer and job.")
        except Exception as e:
            print(f"❌ Error while searching for cutouts: {e}")
            print("Please ensure the database configuration is correct in the setup.")
        
        # input("Done. You may now close this window...")
        refresh_desktop()
    else:
        setup()
        print("No file provided. Please drag an Excel file onto this program.")
