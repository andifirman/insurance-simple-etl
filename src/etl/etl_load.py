# Load step
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



def save_cf_gen(
    df,
    icg_col='ICG',
    dest_folder='data/processed',
    output_name='CF_Gen.xlsx'
):

    os.makedirs(dest_folder, exist_ok=True)
    output_path = os.path.join(dest_folder, output_name)

    # ===============================
    # Column groupings
    # ===============================
    money_columns = [
        'Expected_Premium', 'Expected_Commission', 'Expected_Acquisition',
        'Earned_Premium', 'Earned_Commission', 'Exp_Premium', 'Exp_Commission',
        'Exp_Acquisition', 'Exp_Expense', 'Exp_Claim', 'Exp_RA', 'Exp_NPR',
        'Actual_Premium', 'Actual_Commission', 'Actual_Acquisition'
    ]

    ratio_columns = [
        'Loss_Ratio', 'Risk_Adjustment_Ratio', 'PME_Ratio', 'ULAE_Ratio',
        'Cancellation_Ratio', 'Premium_Refund_Ratio', 'NPR_Ratio',
        'Probability_of_Inforce'
    ]

    integer_columns = ['#Incurred', 'Cohort']

    # kolom tanggal yang mau ditampilkan sebagai 30-Sep-2025
    date_columns = ['Incurred', 'Valuation']

    # ===============================
    # Enforce dtypes
    # ===============================
    for col in money_columns + ratio_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    for col in integer_columns:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(df[col], errors='coerce')
                .fillna(0)
                .astype(int)
            )

    # pastikan kolom tanggal bertipe datetime dulu
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # SEKARANG: ubah tanggal jadi string "dd-MMM-yyyy" untuk tampilan Excel
    for col in date_columns:
        if col in df.columns:
            df[col] = df[col].dt.strftime('%d-%b-%Y')

    # ===============================
    # Write Excel (1 sheet per ICG)
    # ===============================
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book

        # format angka
        format_money = workbook.add_format({
            "num_format": '#,##0;(#,##0)'
        })
        format_ratio = workbook.add_format({
            "num_format": '0.000%'
        })
        format_integer = workbook.add_format({
            "num_format": '0'
        })
        format_default = workbook.add_format({
            "num_format": '0.000'
        })

        for icg_name, df_icg in df.groupby(icg_col):

            # Sheet name handling
            sheet_name = f"CF_{str(icg_name)}"
            sheet_name = sheet_name.replace('/', '_')[:31]

            df_icg.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name
            )

            worksheet = writer.sheets[sheet_name]

            # ===============================
            # Apply column formatting
            # ===============================
            for col in df_icg.columns:
                col_idx = df_icg.columns.get_loc(col)
                col_letter = get_excel_col_letter(col_idx)

                if col in money_columns:
                    worksheet.set_column(
                        f'{col_letter}:{col_letter}',
                        18,
                        format_money
                    )
                elif col in ratio_columns:
                    worksheet.set_column(
                        f'{col_letter}:{col_letter}',
                        15,
                        format_ratio
                    )
                elif col in integer_columns:
                    worksheet.set_column(
                        f'{col_letter}:{col_letter}',
                        12,
                        format_integer
                    )
                else:
                    # termasuk kolom tanggal yang sekarang sudah string
                    worksheet.set_column(
                        f'{col_letter}:{col_letter}',
                        15,
                        format_default
                    )

    # ===============================
    # Style header (existing helper)
    # ===============================
    style_excel_header(output_path)

    return output_path