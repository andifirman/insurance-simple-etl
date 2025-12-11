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
    
    # Kolom-kolom yang memerlukan pemisah ribuan dan pembulatan ke bilangan bulat
    columns_to_format = {
        'Expected Cashflow': ['Premi SOP', 'Premi EOP', 'Claim', 'RA', 'Expense', 'Komisi SOP', 'Komisi EOP'],
        'Actual': ['Premi', 'Komisi']
    }

    # hanya tambahkan if_sheet_exists saat mode='a'
    if file_exists:
        writer_args["if_sheet_exists"] = "replace"

    with pd.ExcelWriter(**writer_args) as writer:

        for sheet_name, df in csm_dict.items():

            # Cek apakah sheet_name adalah Expected Cashflow atau Actual
            if sheet_name in columns_to_format:
                # Kolom-kolom yang harus diformat di sheet ini
                columns_to_process = columns_to_format[sheet_name]

                # Pastikan kolom-kolom yang dipilih sudah menjadi numerik dan format ribuan
                for col in columns_to_process:
                    if col in df.columns:
                        # Mengonversi kolom ke numerik jika belum
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                        
                        # Ganti NaN dan inf dengan 0 sebelum mengonversi ke integer
                        df[col] = df[col].fillna(0)  # Ganti NaN dengan 0
                        df[col] = df[col].replace([float('inf'), -float('inf')], 0)  # Ganti inf dan -inf dengan 0

                        # Membulatkan angka dan memastikan integer (bilangan bulat)
                        df[col] = df[col].round(0).astype(int)

                        # Menambahkan pemisah ribuan
                        df[col] = df[col].apply(lambda x: f"{x:,}")  # Format dengan ribuan separator

            # Jika sheet kosong, lewati penulisan
            if df.empty:
                print(f"Warning: Sheet {sheet_name} is empty. Skipping writing this sheet.")
                continue

            # Jika file sudah ada, coba gabungkan data lama dan data baru
            if file_exists:
                try:
                    # Baca data lama dari sheet yang ada
                    df_old = pd.read_excel(output_path, sheet_name=sheet_name)

                    # Gabungkan df_old dengan df yang baru, misalnya menggunakan pd.concat
                    df = pd.concat([df_old, df], ignore_index=True)

                except ValueError:
                    # Jika sheet tidak ada, maka abaikan dan lanjutkan dengan df baru
                    pass

            # Tulis dataframe ke excel
            df.to_excel(writer, index=False, sheet_name=sheet_name)

    style_excel_header(output_path)  # Terapkan styling header untuk seluruh sheet
    return output_path