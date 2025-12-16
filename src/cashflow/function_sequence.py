import pandas as pd
import os

def generate_incurred_sequence(
    df,
    years_forward: int,
    icg_col='ICG',
    target_col='#Incurred'
):
    """
    Generate #Incurred = 1 .. (years_forward * 4)
    for EACH ICG (stateless, deterministic)
    """

    df = df.copy()
    df[icg_col] = df[icg_col].astype(str).str.strip()

    total_quarters = years_forward * 4

    # Ambil ICG unik
    icg_list = df[icg_col].unique()

    rows = []

    for icg in icg_list:
        base_row = df[df[icg_col] == icg].iloc[0].to_dict()

        for q in range(1, total_quarters + 1):
            r = base_row.copy()
            r[target_col] = q
            rows.append(r)

    return pd.DataFrame(rows)
