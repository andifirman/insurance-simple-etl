# Extract step

import pandas as pd

def load_input_excel(path):
    try:
        xls = pd.ExcelFile(path)
        sheet_names = xls.sheet_names

        print(f'Found {len(sheet_names)} sheets:')
        for s in sheet_names:
            print(' -', s)

        data = {sheet: pd.read_excel(xls, sheet) for sheet in sheet_names}

        print('\nAll sheets loaded successfully.')
        return data

    except Exception as e:
        print(f'Error loading Excel file: {e}')
        return None
