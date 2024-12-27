import pandas as pd
import re
from tkinter import filedialog, messagebox
from talent_manager import update_talent_list, create_permanent_buttons
from globals import globals

def load_excel_file():
    file_path = filedialog.askopenfilename(
        filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")))
    if not file_path:
        return

    globals.talent_cache = {}
    # Le reste du code reste identique, mais utilisez globals.talent_cache au lieu de talent_cache
    # [Garder le code existant mais remplacer talent_cache par globals.talent_cache]

    update_talent_list(globals.talent_cache)
    create_permanent_buttons()