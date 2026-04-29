from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd


@dataclass(frozen=True)
class ReportingFrequency:
    key: str  # 'quarterly' | 'semester' | 'annual'


def _clear_screen() -> None:
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
    except Exception:
        print("\n" * 50)


def prompt_reporting_frequency() -> ReportingFrequency:
    """Single-select reporting frequency.

    Returns:
        ReportingFrequency with key in {'quarterly','semester','annual'}.

    UX: simple terminal input; default is quarterly.
    """

    items = [
        ("Quarterly", "quarterly"),
        ("Semester", "semester"),
        ("Annual", "annual"),
    ]

    if os.name == 'nt':
        try:
            import msvcrt  # type: ignore

            cursor = 0

            def render() -> None:
                _clear_screen()
                print("Pilih frekuensi report yang akan digenerate")
                print("(↑/↓) move | (Enter) OK | (Q/Esc) cancel (default Quarterly)")
                print("-")
                for i, (label, _key) in enumerate(items):
                    pointer = ">" if i == cursor else " "
                    print(f"{pointer} {label}")
                print("-")

            render()
            while True:
                ch = msvcrt.getwch()

                if ch in ('\x00', '\xe0'):
                    ch2 = msvcrt.getwch()
                    if ch2 == 'H':
                        cursor = max(0, cursor - 1)
                        render()
                        continue
                    if ch2 == 'P':
                        cursor = min(len(items) - 1, cursor + 1)
                        render()
                        continue
                    continue

                if ch in ('q', 'Q', '\x1b'):
                    return ReportingFrequency("quarterly")

                if ch in ('\r', '\n'):
                    return ReportingFrequency(items[cursor][1])
        except Exception:
            # fall back to text prompt
            pass

    # fallback (non-Windows / non-interactive)
    options = {"1": "quarterly", "2": "semester", "3": "annual"}
    print("Pilih frekuensi report yang akan digenerate:")
    print("  1) Quarterly (default)")
    print("  2) Semester")
    print("  3) Annual")
    raw = input("Pilihan [1/2/3]: ").strip()
    if raw == "":
        return ReportingFrequency("quarterly")
    key = options.get(raw)
    if key is None:
        raise ValueError("Input tidak valid. Pilih 1, 2, atau 3.")
    return ReportingFrequency(key)


def _bucket_end_dates(dts: pd.Series, frequency_key: str) -> pd.Series:
    dts = pd.to_datetime(dts, errors="coerce")
    if frequency_key == "quarterly":
        return dts

    years = dts.dt.year

    if frequency_key == "annual":
        return pd.to_datetime(
            {
                "year": years,
                "month": 12,
                "day": 31,
            },
            errors="coerce",
        )

    if frequency_key == "semester":
        # H1 ends Jun-30, H2 ends Dec-31
        half = (dts.dt.month > 6).astype(int)  # 0 for H1, 1 for H2
        month = half.map({0: 6, 1: 12})
        day = half.map({0: 30, 1: 31})
        return pd.to_datetime(
            {
                "year": years,
                "month": month,
                "day": day,
            },
            errors="coerce",
        )

    raise ValueError(f"Unknown frequency: {frequency_key}")


def aggregate_cf_for_reporting(cf: pd.DataFrame, frequency_key: str) -> pd.DataFrame:
    """Aggregate CF_Gen to the selected reporting frequency.

    Assumes CF is generated quarterly; aggregation is done for export only.
    """

    if frequency_key == "quarterly":
        return cf

    df = cf.copy()
    original_cols = list(df.columns)

    df["Incurred"] = pd.to_datetime(df["Incurred"], errors="coerce")
    df["Valuation"] = pd.to_datetime(df["Valuation"], errors="coerce")

    df["_Period_End"] = _bucket_end_dates(df["Incurred"], frequency_key)

    money_cols: List[str] = [
        'Expected_Premium', 'Expected_Commission', 'Expected_Acquisition',
        'Earned_Premium', 'Earned_Commission',
        'Exp_Premium', 'Exp_Commission', 'Exp_Acquisition',
        'Exp_Expense', 'Exp_Claim', 'Exp_RA', 'Exp_NPR',
        'Actual_Premium', 'Actual_Commission', 'Actual_Acquisition',
    ]

    ratio_cols: List[str] = [
        'Loss_Ratio', 'Risk_Adjustment_Ratio', 'PME_Ratio', 'ULAE_Ratio',
        'Cancellation_Ratio', 'Premium_Refund_Ratio', 'NPR_Ratio',
        'Probability_of_Inforce_at_BoP', 'Inflation_Ratio', 'Inflation_Factor_at_BoP',
    ]

    group_keys = ["ICG", "_Period_End"]

    for c in money_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in ratio_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Ensure stable ordering so "last" is meaningful
    sort_cols = ["ICG", "_Period_End"]
    if "#Incurred" in df.columns:
        sort_cols.insert(2, "#Incurred")
    sort_cols.append("Incurred")
    df = df.sort_values(sort_cols, kind="mergesort")

    agg: Dict[str, str] = {}
    for col in df.columns:
        if col in group_keys:
            continue
        if col == "Incurred":
            continue
        if col == "#Incurred":
            continue

        if col in money_cols:
            agg[col] = "sum"
        elif col in ratio_cols:
            agg[col] = "last"
        elif col == "Valuation":
            agg[col] = "first"
        else:
            agg[col] = "first"

    out = df.groupby(group_keys, sort=False, dropna=False).agg(agg).reset_index()

    # rebuild Incurred as period end
    out["Incurred"] = out["_Period_End"]
    out = out.drop(columns=["_Period_End"], errors="ignore")

    # rebuild #Incurred per ICG
    out = out.sort_values(["ICG", "Incurred"], kind="mergesort")
    out["#Incurred"] = out.groupby("ICG").cumcount() + 1

    # Match original (quarterly) column order as much as possible
    desired = [c for c in original_cols if c in out.columns]
    remainder = [c for c in out.columns if c not in desired]
    out = out[desired + remainder]

    return out


def aggregate_csm_for_reporting(csm_dict: dict, frequency_key: str) -> dict:
    """Aggregate CSM sheets to the selected reporting frequency for export."""

    if frequency_key == "quarterly":
        return csm_dict

    out = dict(csm_dict)

    def bucket_end(dts: pd.Series) -> pd.Series:
        return _bucket_end_dates(dts, frequency_key)

    # Expected Cashflow
    if "Expected Cashflow" in out:
        df = out["Expected Cashflow"].copy()
        original_cols = list(df.columns)
        df["Incurred"] = pd.to_datetime(df["Incurred"], errors="coerce")
        df["Valuation"] = pd.to_datetime(df["Valuation"], errors="coerce")
        df["_Period_End"] = bucket_end(df["Incurred"])

        money = ['Premi SOP', 'Premi EOP', 'Claim', 'RA', 'Expense', 'Komisi SOP', 'Komisi EOP']
        meta_first = ['Valuation', 'Cohort', 'LOB']

        for c in money:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        df = df.sort_values(["ICG", "_Period_End", "Incurred"], kind="mergesort")
        agg = {c: 'sum' for c in money if c in df.columns}
        for c in meta_first:
            if c in df.columns:
                agg[c] = 'first'

        grouped = df.groupby(["ICG", "_Period_End"], sort=False, dropna=False).agg(agg).reset_index()
        grouped["Incurred"] = grouped["_Period_End"]
        grouped = grouped.drop(columns=["_Period_End"], errors="ignore")

        # format back like existing CSM (mm/dd/yyyy string)
        grouped["Valuation"] = pd.to_datetime(grouped["Valuation"], errors="coerce").dt.strftime("%m/%d/%Y")
        grouped["Incurred"] = pd.to_datetime(grouped["Incurred"], errors="coerce").dt.strftime("%m/%d/%Y")

        desired = [c for c in original_cols if c in grouped.columns]
        remainder = [c for c in grouped.columns if c not in desired]
        out["Expected Cashflow"] = grouped[desired + remainder]

    # Actual
    if "Actual" in out:
        df = out["Actual"].copy()
        original_cols = list(df.columns)
        df["Incurred"] = pd.to_datetime(df["Incurred"], errors="coerce")
        df["_Period_End"] = bucket_end(df["Incurred"])

        money = [
            'Actual Premi SOP', 'Actual Komisi SOP',
            'Actual Premi EOP', 'Actual Komisi EOP',
            'Actual Paid Claim', 'Actual Paid Expense',
            'Investment Income',
        ]
        meta_first = ['Cohort', 'LOB']

        for c in money:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        df = df.sort_values(["ICG", "_Period_End", "Incurred"], kind="mergesort")
        agg = {c: 'sum' for c in money if c in df.columns}
        for c in meta_first:
            if c in df.columns:
                agg[c] = 'first'

        grouped = df.groupby(["ICG", "_Period_End"], sort=False, dropna=False).agg(agg).reset_index()
        grouped["Incurred"] = grouped["_Period_End"]
        grouped = grouped.drop(columns=["_Period_End"], errors="ignore")

        grouped["Incurred"] = pd.to_datetime(grouped["Incurred"], errors="coerce").dt.strftime("%m/%d/%Y")
        desired = [c for c in original_cols if c in grouped.columns]
        remainder = [c for c in grouped.columns if c not in desired]
        out["Actual"] = grouped[desired + remainder]

    # Assumption: one row per ICG (rates are per ICG)
    if "Assumption" in out:
        df = out["Assumption"].copy()
        original_cols = list(df.columns)
        if "ICG" in df.columns:
            df = df.drop_duplicates(subset=["ICG"]).reset_index(drop=True)
        desired = [c for c in original_cols if c in df.columns]
        remainder = [c for c in df.columns if c not in desired]
        out["Assumption"] = df[desired + remainder]

    return out
