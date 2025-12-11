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
    cf_df, 
    cancel_col='Cancellation_Ratio', 
    output_col='Probability_of_Inforce', 
    icg_col='ICG', 
    incurred_col='#Incurred'
):
    # Kolom output
    cf_df[output_col] = 0.0

    # Sort utama
    cf_df = cf_df.sort_values(by=[icg_col, incurred_col]).reset_index(drop=True)

    # Loop tiap ICG
    for icg in cf_df[icg_col].unique():

        # Index asli (di cf_df) untuk ICG tersebut
        idx_list = cf_df.index[cf_df[icg_col] == icg].tolist()

        prob_list = []

        # Loop berdasarkan posisi urut
        for pos, real_idx in enumerate(idx_list):

            row = cf_df.loc[real_idx]

            cancel_ratio = row.get(cancel_col, 0)
            if cancel_ratio is None or pd.isna(cancel_ratio):
                cancel_ratio = 0

            incurred_val = row.get(incurred_col, 0)

            # Logic probability
            if pos == 0:
                # First record
                if incurred_val == 1:
                    prob = 1.0
                else:
                    prob = 1.0 * (1 - cancel_ratio)
            else:
                prev_prob = prob_list[-1]
                if incurred_val == 1:
                    prob = 1.0
                else:
                    prob = prev_prob * (1 - cancel_ratio)

            prob_list.append(prob)

            # Update ke cf_df pakai index asli
            cf_df.loc[real_idx, output_col] = prob

    return cf_df






def generate_exp(cf_df,
                 icg_col='ICG',
                 incurred_col='Incurred',
                 valuation_col='Valuation',
                 expected_premium_col='Expected_Premium',
                 prob_inforce_col='Probability_of_Inforce',
                 premium_refund_col='Premium_Refund_Ratio',
                 cancel_col='Cancellation_Ratio',
                 earned_premium_col='Earned_Premium',
                 output_col='Output_Col'):

    numeric_cols = [
        expected_premium_col,
        prob_inforce_col,
        premium_refund_col,
        cancel_col,
        earned_premium_col
    ]

    for c in numeric_cols:
        cf_df[c] = pd.to_numeric(cf_df[c], errors='coerce').fillna(0)

    cf_df[incurred_col] = pd.to_datetime(cf_df[incurred_col], errors='coerce')
    cf_df[valuation_col] = pd.to_datetime(cf_df[valuation_col], errors='coerce')

    cf_df = cf_df.sort_values(by=[icg_col, incurred_col]).reset_index(drop=True)

    cf_df[output_col] = 0.0  
    for icg in cf_df[icg_col].unique():
        mask = cf_df[icg_col] == icg
        sub_df = cf_df.loc[mask].copy()
        cum_earned = sub_df[earned_premium_col][::-1].cumsum()[::-1] 

        for idx, row in sub_df.iterrows():
            if row[incurred_col] <= row[valuation_col]:
                cf_df.loc[idx, output_col] = row[expected_premium_col]
            else:
                cf_df.loc[idx, output_col] = - (
                    row[prob_inforce_col] * row[premium_refund_col] * row[cancel_col] * cum_earned.loc[idx]
                )

    cf_df[incurred_col] = cf_df[incurred_col].dt.strftime('%d-%b-%Y')
    cf_df[valuation_col] = cf_df[valuation_col].dt.strftime('%d-%b-%Y')

    return cf_df
  

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
