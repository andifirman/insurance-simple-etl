import tkinter as tk
from tkinter import filedialog
import shutil
import os

def get_input_data(dest_folder='data/raw'):
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title='Select Input Excel File',
        filetypes=[('Excel Files', '*.xlsx *.xls')]
    )

    if not file_path:
        print('No file selected.')
        return None

    os.makedirs(dest_folder, exist_ok=True)

    dest_path = os.path.join(dest_folder, os.path.basename(file_path))
    shutil.copy(file_path, dest_path)

    print(f'File saved to: {dest_path}')
    return dest_path
