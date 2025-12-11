import tkinter as tk
from tkinter import simpledialog

import pandas as pd

from datetime import datetime
from pandas.tseries.offsets import MonthEnd


def get_valuation_date(df_valuation):
    for _, row in df_valuation.iterrows():
        label = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
        if "VALUATION DATE" in label:
            raw_date = row.iloc[1]
            return pd.to_datetime(raw_date, dayfirst=True).normalize()

    # fallback: first non-null in second column
    if df_valuation.shape[1] >= 2:
        col_b = df_valuation.iloc[:, 1].dropna()
        if not col_b.empty:
            return pd.to_datetime(col_b.iloc[0], dayfirst=True).normalize()

    raise ValueError("Valuation date not found in ValuationDate sheet")



def generate_incurred_date(df, valuation_dt=None,
                           valuation_col='Valuation',
                           seq_col='#Incurred',
                           target_col='Incurred'):

    df = df.copy()

    def _calc(row):
        # choose valuation source
        val = row.get(valuation_col, valuation_dt)
        if pd.isna(val):
            return pd.NaT
        val = pd.to_datetime(val)
        base_year = val.year - 1
        base_date = pd.Timestamp(year=base_year, month=12, day=31)
        months_to_add = int(row.get(seq_col, 1)) * 3
        shifted = base_date + pd.DateOffset(months=months_to_add)
        # EOMONTH -> last day of that month
        return (shifted + MonthEnd(0)).normalize()

    df[target_col] = df.apply(_calc, axis=1)
    return df