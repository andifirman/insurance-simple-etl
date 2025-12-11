import pandas as pd

def parse_icg_master(icg_value: str):
    try:
        if not isinstance(icg_value, str):
            return None, None, None

        # Split pertama berdasarkan '#'
        parts = icg_value.split('#')

        if len(parts) < 3:
            return None, None, None

        cohort = parts[1].strip()
        cohort = int(float(cohort))

        # Bagian kontrak dan portfolio
        contract_portfolio = parts[2].split('-')
        if len(contract_portfolio) != 2:
            return None, None, None

        contract = contract_portfolio[0].strip()
        portfolio = portfolio = f"{contract}-{contract_portfolio[1].strip()}"

        return cohort, contract, portfolio

    except Exception:
        return None, None, None


def join_icg_master(icg_value: str, master_df: pd.DataFrame):
    try:
        row = master_df.loc[master_df['ICG'] == icg_value]

        if row.empty:
            return None, None, None

        return (
            str(row['Cohort'].iloc[0]),
            str(row['Contract'].iloc[0]),
            str(row['Portfolio'].iloc[0])
        )

    except Exception:
        return None, None, None


def extract_icg_values(icg_value: str, master_df: pd.DataFrame = None):
    cohort, contract, portfolio = parse_icg_master(icg_value)

    if cohort and contract and portfolio:
        return cohort, contract, portfolio

    # --- 2. Fallback ke master (jika tersedia) ---
    if master_df is not None:
        return join_icg_master(icg_value, master_df)

    return None, None, None


# Expected Premium
