import pandas as pd

def generate_join_value(lookup_value, df, lookup_col, return_col):
    try:
        match = df.loc[df[lookup_col] == lookup_value, return_col]

        if match.empty:
            return None 
        return match.iloc[0]

    except Exception as e:
        print(f"[Data Not Found] {e}")
        return None


def generate_sum_by(df, sum_col, return_formatted=False, **criteria):
    try:
        filtered = df.copy()

        # Apply semua criteria sebagai filter
        for col, val in criteria.items():
            filtered = filtered[filtered[col] == val]

        total = filtered[sum_col].sum()

        if not return_formatted:
            return abs(total)

        # Jika return_formatted True, format hasilnya
        return total

    except Exception as e:
        print(f"[generate_sum_by ERROR] {e}")
        return 0
      

def generate_npr_ratio(cf_df, assumptions_df, sum_col='NPR_Ratio', contract_col='Contract', contract_val='RC', icg_col='ICG'):
    def lookup_npr(row):
        if row[contract_col] == contract_val:
            try:
                # XLOOKUP: cari ICG di assumptions, ambil kolom NPR_Ratio
                value = assumptions_df.loc[assumptions_df[icg_col] == row[icg_col], sum_col].values
                return value[0] if len(value) > 0 else 0
            except Exception:
                return 0
        else:
            return 0

    cf_df['NPR_Ratio'] = cf_df.apply(lookup_npr, axis=1)
    return cf_df


def generate_probability_of_inforce(
    df,
    cancel_col='Cancellation_Ratio',
    output_col='Probability_of_Inforce',
    icg_col='ICG',
    incurred_col='#Incurred'
):
    df = df.copy()

    df = df.sort_values([icg_col, incurred_col])

    df[output_col] = 1.0

    for icg, g in df.groupby(icg_col):
        probs = []
        for i, row in g.iterrows():
            if not probs:
                probs.append(1.0)
            else:
                probs.append(probs[-1] * (1 - row[cancel_col]))

        df.loc[g.index, output_col] = probs

    return df


# ================================================== #
def exp_premium_formula_row(
    incurred,
    valuation,
    expected_premium,
    probability_of_inforce,
    premium_refund_ratio,
    cancellation_ratio,
    earned_premium_sum_range,
):
    if pd.isna(incurred) or pd.isna(valuation):
        return 0.0

    if incurred <= valuation:
        return expected_premium

    return -(
        probability_of_inforce
        * premium_refund_ratio
        * cancellation_ratio
        * earned_premium_sum_range
    )

def build_earned_sum_per_icg(earned_cf, icg_col='ICG', earned_premium_col='Earned_Premium'):
    tmp = earned_cf.copy()
    tmp[earned_premium_col] = pd.to_numeric(tmp[earned_premium_col], errors='coerce').fillna(0.0)
    
    return tmp.groupby(icg_col)[earned_premium_col].sum()


def generate_exp_premium(
    df,
    icg_col,
    incurred_col,
    valuation_col,
    expected_col,
    prob_inforce_col,
    premium_refund_col,
    cancel_col,
    earned_sum_per_icg,  
    output_col,
    earned_premium_col='Earned_Premium',
):
    df = df.copy()

    df[incurred_col] = pd.to_datetime(df[incurred_col], errors="coerce")
    df[valuation_col] = pd.to_datetime(df[valuation_col], errors="coerce")

    for c in [expected_col, prob_inforce_col, premium_refund_col, cancel_col, earned_premium_col]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    df[output_col] = 0.0

    for icg, g in df.groupby(icg_col, sort=False):
        # total SUM(K2:K$105) untuk ICG ini (dari helper build_earned_sum_per_icg)
        total_earned = float(earned_sum_per_icg.get(icg, 0.0))

        # urutkan baris sama seperti Excel (kalau ada #Incurred)
        sort_cols = []
        if '#Incurred' in g.columns:
            sort_cols.append('#Incurred')
        sort_cols.append(incurred_col)
        g = g.sort_values(sort_cols, kind="mergesort")

        # prefix sum Earned_Premium: SUM(K2:K_(pos-1))
        earned_values = g[earned_premium_col].to_numpy()
        prefix_sum = 0.0

        for pos, (idx, row) in enumerate(g.iterrows(), start=1):
            # pos = 1,2,3,... seiring baris CF
            # Excel: baris pos pakai SUM(K_pos:K$105)
            # Python: total_earned = SUM(K2:K$105)
            #         prefix_sum  = SUM(K2:K_(pos-1))
            earned_sum_range = total_earned - prefix_sum

            df.loc[idx, output_col] = exp_premium_formula_row(
                incurred=row[incurred_col],
                valuation=row[valuation_col],
                expected_premium=row[expected_col],
                probability_of_inforce=row[prob_inforce_col],
                premium_refund_ratio=row[premium_refund_col],
                cancellation_ratio=row[cancel_col],
                earned_premium_sum_range=earned_sum_range,
            )

            # update prefix untuk baris berikutnya
            prefix_sum += earned_values[pos - 1]

    return df

def generate_exp_commission(
    df,
    icg_col,
    incurred_col,
    valuation_col,
    expected_col,
    prob_inforce_col,
    premium_refund_col,
    cancel_col,
    earned_sum_per_icg,  
    output_col,
    earned_premium_col='Earned_Commission',
):
    df = df.copy()

    df[incurred_col] = pd.to_datetime(df[incurred_col], errors="coerce")
    df[valuation_col] = pd.to_datetime(df[valuation_col], errors="coerce")

    for c in [expected_col, prob_inforce_col, premium_refund_col, cancel_col, earned_premium_col]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    df[output_col] = 0.0

    for icg, g in df.groupby(icg_col, sort=False):
        # total SUM(K2:K$105) untuk ICG ini (dari helper build_earned_sum_per_icg)
        total_earned = float(earned_sum_per_icg.get(icg, 0.0))

        # urutkan baris sama seperti Excel (kalau ada #Incurred)
        sort_cols = []
        if '#Incurred' in g.columns:
            sort_cols.append('#Incurred')
        sort_cols.append(incurred_col)
        g = g.sort_values(sort_cols, kind="mergesort")

        # prefix sum Earned_Premium: SUM(K2:K_(pos-1))
        earned_values = g[earned_premium_col].to_numpy()
        prefix_sum = 0.0

        for pos, (idx, row) in enumerate(g.iterrows(), start=1):
            # pos = 1,2,3,... seiring baris CF
            # Excel: baris pos pakai SUM(K_pos:K$105)
            # Python: total_earned = SUM(K2:K$105)
            #         prefix_sum  = SUM(K2:K_(pos-1))
            earned_sum_range = total_earned - prefix_sum

            df.loc[idx, output_col] = exp_premium_formula_row(
                incurred=row[incurred_col],
                valuation=row[valuation_col],
                expected_premium=row[expected_col],
                probability_of_inforce=row[prob_inforce_col],
                premium_refund_ratio=row[premium_refund_col],
                cancellation_ratio=row[cancel_col],
                earned_premium_sum_range=earned_sum_range,
            )

            # update prefix untuk baris berikutnya
            prefix_sum += earned_values[pos - 1]

    return df

  

def generate_expense_claim(cf_df,
                               prob_inforce_col='Probability_of_Inforce',
                               earned_premium_col='Earned_Premium',
                               cancel_col='Cancellation_Ratio',
                               ratio_col=None,
                               exp_claim_col=None,
                               output_expense_col='Exp_Expense',
                               output_claim_col='Exp_Claim'):
  
    # Pastikan semua kolom numerik
    numeric_cols = [
        prob_inforce_col,
        earned_premium_col,
        cancel_col,
        ratio_col, 
        exp_claim_col  
    ]
    for c in numeric_cols:
        cf_df[c] = pd.to_numeric(cf_df[c], errors='coerce').fillna(0)

    # Exp_Claim
    cf_df[output_claim_col] = (
        -cf_df[prob_inforce_col] * cf_df[earned_premium_col] * (1 - cf_df[cancel_col]) * cf_df['Loss_Ratio'] 
    )

    # Exp_Expense
    cf_df[output_expense_col] = (
        -cf_df[prob_inforce_col] * cf_df[earned_premium_col] * (1 - cf_df[cancel_col]) * cf_df[ratio_col] 
        + cf_df[exp_claim_col] * cf_df['ULAE_Ratio'] 
    )

    return cf_df
  
def generate_actual(cf_df,
                    actual_cf_df,
                    icg_col='ICG',
                    incurred_col='Incurred',
                    output_col='Output_Col',
                    target_col='Target_Column'):

    # Pastikan kolom yang diperlukan ada di actual_cf_df
    if icg_col not in actual_cf_df.columns or incurred_col not in actual_cf_df.columns or target_col not in actual_cf_df.columns:
        raise ValueError(f"Kolom yang diperlukan tidak ada di Actual_CF: {icg_col}, {incurred_col}, {target_col}")
    
    # Lakukan SUMIFS untuk setiap baris berdasarkan kondisi ICG dan Incurred
    cf_df[output_col] = cf_df.apply(
        lambda row: actual_cf_df[
            (actual_cf_df[icg_col] == row[icg_col]) & 
            (actual_cf_df[incurred_col] == row[incurred_col])
        ][target_col].sum(), axis=1
    )

    return cf_df


def generate_exp_claim_inflation(
    cf_df,
    claim_inflation_rate,
    incurred_order_col='#Incurred',
    base_claim_col='Exp_Claim',
    output_col='Exp_Claim_Inflation',
):

    cf_df = cf_df.copy()

    # Validasi kolom
    missing = [
        col
        for col in [incurred_order_col, base_claim_col]
        if col not in cf_df.columns
    ]
    if missing:
        raise KeyError(f"Kolom berikut tidak ditemukan di CF_Gen: {missing}")

    # Pastikan numeric
    incurred_order = pd.to_numeric(cf_df[incurred_order_col], errors="coerce")
    base_claim = pd.to_numeric(cf_df[base_claim_col], errors="coerce")

    # Faktor inflasi: (1 + r)^(#Incurred - 1)
    factor = (1.0 + float(claim_inflation_rate)) ** (incurred_order - 1)
    factor = factor.fillna(1.0)

    cf_df[output_col] = base_claim * factor

    return cf_df
  

def generate_exp_expense_inflation(
    cf_df,
    expense_inflation_rate,
    incurred_order_col='#Incurred',
    base_claim_col='Exp_Expense',
    output_col='Exp_Expense_Inflation',
):

    cf_df = cf_df.copy()

    # Validasi kolom
    missing = [
        col
        for col in [incurred_order_col, base_claim_col]
        if col not in cf_df.columns
    ]
    if missing:
        raise KeyError(f"Kolom berikut tidak ditemukan di CF_Gen: {missing}")

    # Pastikan numeric
    incurred_order = pd.to_numeric(cf_df[incurred_order_col], errors="coerce")
    base_claim = pd.to_numeric(cf_df[base_claim_col], errors="coerce")

    # Faktor inflasi: (1 + r)^(#Incurred - 1)
    factor = (1.0 + float(expense_inflation_rate)) ** (incurred_order - 1)
    factor = factor.fillna(1.0)

    cf_df[output_col] = base_claim * factor

    return cf_df
