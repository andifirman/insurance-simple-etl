import pandas as pd
import os

def generate_incurred_sequence(df_new, output_path=None, icg_col='ICG', target_col='#Incurred'):

    # --- Kalau belum ada file → semua #Incurred = 1 ---
    if not os.path.exists(output_path):
        df_new[target_col] = 1
        return df_new

    # --- Load existing history ---
    df_old = pd.read_excel(output_path, sheet_name="CF")

    # Gabungkan data lama + baru
    df_all = pd.concat([df_old, df_new], ignore_index=True)

    # Hitung sequence berdasarkan ICG saja
    df_all[target_col] = df_all.groupby([icg_col]).cumcount() + 1

    #  Ambil hanya baris baru untuk dikembalikan
    df_new[target_col] = df_all.tail(len(df_new))[target_col].values

    return df_new

# def generate_incurred_sequence(
#     df_new,
#     output_path=None,
#     icg_col='ICG',
#     valuation_col='Valuation',
#     target_col='#Incurred'
# ):
#     df_new = df_new.copy()
#     df_new[icg_col] = df_new[icg_col].astype(str).str.strip()

#     # add marker to identify new rows after concat
#     df_new["_is_new"] = True

#     # if no existing file
#     if not output_path or not os.path.exists(output_path):
#         df_new[target_col] = 1
#         return df_new.drop(columns=["_is_new"])

#     # load old data
#     try:
#         df_old = pd.read_excel(output_path, sheet_name="CF")
#     except Exception:
#         df_new[target_col] = 1
#         return df_new.drop(columns=["_is_new"])

#     # normalize
#     df_old = df_old.copy()
#     df_old[icg_col] = df_old[icg_col].astype(str).str.strip()
#     df_old["_is_new"] = False

#     # ensure date
#     if valuation_col in df_old.columns:
#         df_old[valuation_col] = pd.to_datetime(df_old[valuation_col], errors='coerce')
#     df_new[valuation_col] = pd.to_datetime(df_new[valuation_col], errors='coerce')

#     # combine but keep track which rows are new
#     df_all = pd.concat([df_old, df_new], ignore_index=True)

#     # sort by ICG then valuation
#     if valuation_col in df_all.columns:
#         df_all = df_all.sort_values([icg_col, valuation_col, "_is_new"]).reset_index(drop=True)
#     else:
#         df_all = df_all.sort_values([icg_col, "_is_new"]).reset_index(drop=True)

#     # assign fresh incurred sequence for all rows
#     df_all[target_col] = df_all.groupby(icg_col).cumcount() + 1

#     # extract only NEW rows with their updated incurred
#     df_new_final = df_all[df_all["_is_new"]].copy()
#     df_new_final = df_new_final.sort_index()  # keep original new row order
#     df_new_final = df_new_final.drop(columns=["_is_new"])

#     return df_new_final