# Transform step
import pandas as pd
import numpy as np

from src.utils.helper_output_structure import *
from src.cashflow.function_icg_parser import *

from src.utils.helper_output_format import *

from src.cashflow.function_sequence import generate_incurred_sequence
from src.cashflow.function_date import *
from src.cashflow.function_cf_generator import *


def build_cf_gen(all_sheets_dict, years_forward: int, claim_inflation_rate: float, expense_inflation_rate: float):
    """
    Build Cashflow Projection (CF_Gen)
    years_forward = jumlah tahun ke depan (user input)
    """

    # =============================
    # 1. LOAD TEMPLATE + VALIDATION
    # =============================
    cf = create_cf_gen_template()

    # --- Load valuation ---
    df_valuation = all_sheets_dict.get('ValuationDate')
    if df_valuation is None:
        raise ValueError("Sheet 'ValuationDate' tidak ditemukan.")

    valuation_date = get_valuation_date(df_valuation)

    # --- Load ICG ---
    df_icg = all_sheets_dict.get('ICG')
    if df_icg is None:
        raise ValueError("Sheet 'ICG' tidak ditemukan.")

    cf['ICG'] = df_icg['ICG']
    cf['Valuation'] = valuation_date

    # =============================
    # 2. PARSE ICG
    # =============================
    parsed = cf['ICG'].apply(extract_icg_values)

    cf['Cohort'] = (
        parsed.apply(lambda x: x[0])
              .pipe(pd.to_numeric, errors='coerce')
              .astype('Int64')
              .astype(str)
    )
    cf['Contract']  = parsed.apply(lambda x: x[1])
    cf['Portfolio'] = parsed.apply(lambda x: x[2])

    # =============================
    # 3. INCURRED SEQUENCE (NEW LOGIC)
    # =============================
    cf = generate_incurred_sequence(
        cf,
        years_forward=years_forward,
        icg_col='ICG',
        target_col='#Incurred'
    )

    cf = generate_incurred_date(
        cf,
        valuation_dt=cf['Valuation'].iloc[0],
        incurred_col='#Incurred',
        output_col='Incurred'
    )

    # =============================
    # 4. EXPECTED CF
    # =============================
    expected_cf = all_sheets_dict.get('Expected_CF')
    if expected_cf is None:
        raise ValueError("Sheet 'Expected_CF' tidak ditemukan.")

    def apply_expected(row, col):
        return generate_sum_by(
            expected_cf,
            sum_col=col,
            return_formatted=True,
            ICG=row['ICG'],
            Incurred=row['Incurred']
        )

    cf['Expected_Premium'] = cf.apply(
        lambda r: apply_expected(r, 'Expected_Premium'), axis=1
    )

    cf['Expected_Commission'] = cf.apply(
        lambda r: apply_expected(r, 'Expected_Commission'), axis=1
    )

    cf['Expected_Acquisition'] = cf.apply(
        lambda r: apply_expected(r, 'Expected_Acquisition')
        if r['Contract'] == 'IC' else 0,
        axis=1
    )

    # =============================
    # 5. EARNED CF
    # =============================
    earned_cf = all_sheets_dict.get('Earned')
    if earned_cf is None:
        raise ValueError("Sheet 'Earned' tidak ditemukan.")

    def apply_earned(row, col):
        return generate_sum_by(
            earned_cf,
            sum_col=col,
            return_formatted=True,
            ICG=row['ICG'],
            Incurred=row['Incurred']
        )

    cf['Earned_Premium'] = cf.apply(
        lambda r: apply_earned(r, 'Earned_Premium'), axis=1
    )

    cf['Earned_Commission'] = cf.apply(
        lambda r: apply_earned(r, 'Earned_Commission'), axis=1
    )

		# SUM(K2:K$105) per ICG untuk Premium & Commission
    earned_sum_per_icg_prem = build_earned_sum_per_icg(
        earned_cf,
        icg_col='ICG',
        earned_premium_col='Earned_Premium',
    )
    earned_sum_per_icg_comm = build_earned_sum_per_icg(
        earned_cf,
        icg_col='ICG',
        earned_premium_col='Earned_Commission',
    )

    # =============================
    # 6. ASSUMPTIONS JOIN
    # =============================
    assumptions_df = all_sheets_dict.get('Assumptions')
    if assumptions_df is None:
        raise ValueError("Sheet 'Assumptions' tidak ditemukan.")

    assumption_cols = [
        'Loss_Ratio', 'Risk_Adjustment_Ratio',
        'PME_Ratio', 'ULAE_Ratio',
        'Cancellation_Ratio', 'Premium_Refund_Ratio'
    ]

    for col in assumption_cols:
        cf[col] = cf['ICG'].apply(
            lambda x: generate_join_value(x, assumptions_df, 'ICG', col)
        )

    # =============================
    # 7. PROBABILITY OF INFORCE
    # =============================
    cf = generate_probability_of_inforce(
        cf,
        cancel_col='Cancellation_Ratio',
        output_col='Probability_of_Inforce',
        icg_col='ICG',
        incurred_col='#Incurred'
    )

    # =============================
    # 8. NPR RATIO
    # =============================
    cf = generate_npr_ratio(
        cf,
        assumptions_df=assumptions_df,
        sum_col='NPR_Ratio',
        contract_col='Contract',
        contract_val='RC',
        icg_col='ICG'
    )

    # =============================
    # 9. EXPECTED RESULTS
    # =============================
    cf = generate_exp_premium(
        cf,
        icg_col='ICG',
        incurred_col='Incurred',
        valuation_col='Valuation',
        expected_col='Expected_Premium',
        prob_inforce_col='Probability_of_Inforce',
        premium_refund_col='Premium_Refund_Ratio',
        cancel_col='Cancellation_Ratio',
        earned_sum_per_icg=earned_sum_per_icg_prem,
        output_col='Exp_Premium',
    )

    cf = generate_exp_commission(
				cf,
				icg_col='ICG',
				incurred_col='Incurred',
				valuation_col='Valuation',
				expected_col='Expected_Commission',
				prob_inforce_col='Probability_of_Inforce',
				premium_refund_col='Premium_Refund_Ratio',
				cancel_col='Cancellation_Ratio',
				earned_sum_per_icg=earned_sum_per_icg_comm,
				output_col='Exp_Commission',
				earned_premium_col='Earned_Commission',
		)
    

    cf['Exp_Acquisition'] = cf['Expected_Acquisition']

    # =============================
    # 10. EXPENSE + CLAIM
    # =============================
    cf = generate_expense_claim(
        cf,
        prob_inforce_col='Probability_of_Inforce',
        earned_premium_col='Earned_Premium',
        cancel_col='Cancellation_Ratio',
        ratio_col='PME_Ratio',
        exp_claim_col='Exp_Claim',
        output_expense_col='Exp_Expense',
        output_claim_col='Exp_Claim'
    )

    cf['Exp_RA'] = cf['Exp_Claim'] * cf['Risk_Adjustment_Ratio']

    cf['Exp_NPR'] = np.where(
        cf['Contract'] == 'RC',
        -cf['NPR_Ratio'] * cf['Exp_Claim'],
        0
    )

    # =============================
    # 11. ACTUAL CF
    # =============================
    actual_cf = all_sheets_dict.get('Actual_CF')
    if actual_cf is None:
        raise ValueError("Sheet 'Actual_CF' tidak ditemukan.")

    actual_cols = [
        ('Actual_Premium', 'Actual_Premium'),
        ('Actual_Commission', 'Actual_Commission'),
        ('Actual_Acquisition', 'Actual_Acquisition'),
    ]

    for output_col, target_col in actual_cols:
        cf = generate_actual(
            cf,
            actual_cf,
            icg_col='ICG',
            incurred_col='Incurred',
            output_col=output_col,
            target_col=target_col
        )
        
  	# =============================
    # 12. CLAIM INFLATION
    # =============================
    cf = generate_exp_claim_inflation(
        cf,
        claim_inflation_rate=claim_inflation_rate,
        incurred_order_col='#Incurred',
        base_claim_col='Exp_Claim',
        output_col='Exp_Claim_Inflation',
    )
    
    # =============================
    # 13. EXPENSE INFLATION
    # =============================
    cf = generate_exp_expense_inflation(
        cf,
        expense_inflation_rate=expense_inflation_rate,
        incurred_order_col='#Incurred',
        base_claim_col='Exp_Expense',
        output_col='Exp_Expense_Inflation',
    )

    return cf
