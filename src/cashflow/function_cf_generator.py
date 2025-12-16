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





def generate_exp(cf_df,
                 icg_col='ICG',
                 incurred_col='Incurred',
                 valuation_col='Valuation',
                 expected_premium_col='Expected_Premium',
                 prob_inforce_col='Probability_of_Inforce',
                 premium_refund_col='Premium_Refund_Ratio',
                 cancel_col='Cancellation_Ratio',
                 earned_premium_col='Earned_Premium',
                 output_col='Exp_Premium'):
		
    df = cf_df.copy()
      
    numeric_cols = [
        expected_premium_col,
        prob_inforce_col,
        premium_refund_col,
        cancel_col,
        earned_premium_col
    ]

    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    df[incurred_col] = pd.to_datetime(df[incurred_col], errors='coerce')
    df[valuation_col] = pd.to_datetime(df[valuation_col], errors='coerce')

    df = df.sort_values(by=[icg_col, incurred_col]).reset_index(drop=True)

    df[output_col] = 0.0  
    for icg, g in df.groupby(icg_col):
        # cumulative earned premium dari belakang (future exposure)
        cum_earned = g[earned_premium_col][::-1].cumsum()[::-1]

        for idx in g.index:
            row = df.loc[idx]

            # PAST CF (<= valuation)
            if row[incurred_col] <= row[valuation_col]:
                df.loc[idx, output_col] = row[expected_premium_col]

            # FUTURE CF (> valuation)
            else:
                df.loc[idx, output_col] = - (
                    row[prob_inforce_col]
                    * row[premium_refund_col]
                    * row[cancel_col]
                    * cum_earned.loc[idx]
                )

    df[incurred_col] = df[incurred_col].dt.strftime('%d-%b-%Y')
    df[valuation_col] = df[valuation_col].dt.strftime('%d-%b-%Y')

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
