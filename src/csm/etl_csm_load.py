import pandas as pd
import os

from src.utils.helper_output_format import style_excel_header
from src.utils.helper_output_structure import *



def get_excel_col_letter(col_idx):
  # Function to get All Column
    letters = ""
    while col_idx >= 0:
        letters = chr(col_idx % 26 + 65) + letters
        col_idx = col_idx // 26 - 1
    return letters



def save_csm_gen(csm_dict, output_path):
    file_exists = os.path.exists(output_path)

    # build args dinamis
    writer_args = {
        "path": output_path,
        "engine": "openpyxl",  # menggunakan openpyxl untuk penulisan
        "mode": "a" if file_exists else "w"
    }
    
    money_columns_by_sheet = {
        'Expected Cashflow': ['Premi SOP', 'Premi EOP', 'Claim', 'RA', 'Expense', 'Komisi SOP', 'Komisi EOP'],
        'Actual': [
            'Actual Premi SOP',
            'Actual Komisi SOP',
            'Actual Premi EOP',
            'Actual Komisi EOP',
            'Actual Paid Claim',
            'Actual Paid Expense',
            'Investment Income',
        ]
    }

    # hanya tambahkan if_sheet_exists saat mode='a'
    if file_exists:
        writer_args["if_sheet_exists"] = "replace"

    with pd.ExcelWriter(**writer_args) as writer:

        for sheet_name, df in csm_dict.items():
            df_to_write = df.copy()

            if sheet_name in money_columns_by_sheet:
                for col in money_columns_by_sheet[sheet_name]:
                    if col in df_to_write.columns:
                        df_to_write[col] = (
                            pd.to_numeric(df_to_write[col], errors='coerce')
                            .replace([float('inf'), -float('inf')], 0)
                        )


            # Tulis dataframe ke excel
            df_to_write.to_excel(writer, index=False, sheet_name=sheet_name)

            worksheet = writer.sheets[sheet_name]

            for col_idx, col_name in enumerate(df_to_write.columns, start=1):
                col_letter = get_excel_col_letter(col_idx - 1)

                if sheet_name in money_columns_by_sheet and col_name in money_columns_by_sheet[sheet_name]:
                    worksheet.column_dimensions[col_letter].width = 18
                    for row_idx in range(2, len(df_to_write) + 2):
                        worksheet.cell(row=row_idx, column=col_idx).number_format = '#,##0;(#,##0)'
                else:
                    worksheet.column_dimensions[col_letter].width = 15

    style_excel_header(output_path)  # Terapkan styling header untuk seluruh sheet
    return output_path
