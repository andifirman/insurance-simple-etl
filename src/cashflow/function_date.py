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


def generate_incurred_date(
    df,
    valuation_dt,
    incurred_col='Incurred',
    output_col='Incurred_Date'
):
    """
    Excel equivalent:
    =EOMONTH(DATE(YEAR(Valuation)-1,12,31), Incurred*3)
    """

    base_date = pd.Timestamp(
        year=valuation_dt.year - 1,
        month=12,
        day=31
    )

    df[output_col] = df[incurred_col].apply(
        lambda x: base_date + MonthEnd(x * 3)
    )

    return df
