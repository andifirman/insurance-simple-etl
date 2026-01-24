# Transform step
import pandas as pd
import numpy as np
import math

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
    # 2. PARSE / MAP ICG METADATA
    # =============================
    # Newer input format provides explicit columns in sheet 'ICG'
    # (Cohort, Portfolio, Contract). Prefer these when present.
    if {'Cohort', 'Portfolio', 'Contract'}.issubset(set(df_icg.columns)):
        cf['Cohort'] = (
            pd.to_numeric(df_icg['Cohort'], errors='coerce')
            .astype('Int64')
        )
        cf['Contract'] = df_icg['Contract']
        cf['Portfolio'] = df_icg['Portfolio']
    else:
        parsed = cf['ICG'].apply(extract_icg_values)
        cf['Cohort'] = (
            parsed.apply(lambda x: x[0])
            .pipe(pd.to_numeric, errors='coerce')
            .astype('Int64')
        )
        cf['Contract'] = parsed.apply(lambda x: x[1])
        cf['Portfolio'] = parsed.apply(lambda x: x[2])

    # =============================
    # 3. INCURRED SEQUENCE (NEW LOGIC)
    # =============================
    # Excel templates often keep a fixed number of quarter rows (e.g., up to row 129).
    # To keep SUM(Kx:K$end)/SUM(Lx:L$end) parity, ensure the projection horizon
    # covers all quarters present in the Earned sheet.
    earned_for_horizon = all_sheets_dict.get('Earned')
    if earned_for_horizon is not None and 'Incurred' in earned_for_horizon.columns:
        incurred_dt = pd.to_datetime(earned_for_horizon['Incurred'], errors='coerce')
        base_date = pd.Timestamp(year=valuation_date.year - 1, month=12, day=31)
        base_period = base_date.to_period('Q')
        incurred_period = incurred_dt.dt.to_period('Q')

        # quarter distance: Q1 2023 (ordinal) - Q4 2022 (ordinal) = 1
        quarter_diff = incurred_period.apply(
            lambda p: (p.ordinal - base_period.ordinal) if pd.notna(p) else None
        )
        max_q = pd.to_numeric(quarter_diff, errors='coerce').max()
        if pd.notna(max_q):
            max_quarters_needed = int(max(0, max_q))
            years_forward = max(int(years_forward), int(math.ceil(max_quarters_needed / 4)))

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


    # SUM(K2:K$end) / SUM(L2:L$end) per ICG untuk Premium & Commission
    # IMPORTANT: Excel refers to CF_Gen columns, not the raw Earned sheet.
    earned_sum_per_icg_prem = (
        cf.assign(Earned_Premium=pd.to_numeric(cf['Earned_Premium'], errors='coerce').fillna(0.0))
          .groupby('ICG', sort=False)['Earned_Premium']
          .sum()
    )
    earned_sum_per_icg_comm = (
        cf.assign(Earned_Commission=pd.to_numeric(cf['Earned_Commission'], errors='coerce').fillna(0.0))
          .groupby('ICG', sort=False)['Earned_Commission']
          .sum()
    )

    # =============================
    # 6. ASSUMPTIONS JOIN
    # =============================
    assumptions_df = all_sheets_dict.get('Assumptions')
    if assumptions_df is None:
        raise ValueError("Sheet 'Assumptions' tidak ditemukan.")

    # Canonical CF_Gen columns we want to populate from Assumptions.
    # Some source headers may differ in the input workbook.
    assumption_cols = [
        'Loss_Ratio', 'Risk_Adjustment_Ratio',
        'PME_Ratio', 'ULAE_Ratio',
        'Cancellation_Ratio', 'Premium_Refund_Ratio'
    ]

    # Vectorized XLOOKUP(ICG -> assumption col)
    assumptions_df = assumptions_df.copy()
    if 'ICG' not in assumptions_df.columns:
        raise ValueError("Kolom 'ICG' tidak ditemukan di sheet Assumptions.")

    assumptions_idx = assumptions_df.drop_duplicates(subset=['ICG']).set_index('ICG')

    source_candidates = {
        # Excel header in the updated template
        'Cancellation_Ratio': ['Cancellation_Ratio', 'Cancellation per year (%)'],
        'Premium_Refund_Ratio': ['Premium_Refund_Ratio'],
        'Loss_Ratio': ['Loss_Ratio'],
        'Risk_Adjustment_Ratio': ['Risk_Adjustment_Ratio'],
        'PME_Ratio': ['PME_Ratio'],
        'ULAE_Ratio': ['ULAE_Ratio'],
    }

    for cf_col in assumption_cols:
        candidates = source_candidates.get(cf_col, [cf_col])
        src_col = next((c for c in candidates if c in assumptions_idx.columns), None)
        if src_col is None:
            cf[cf_col] = 0.0
        else:
            cf[cf_col] = cf['ICG'].map(assumptions_idx[src_col])

    # =============================
    # 6B. APPLY UPDATED EXCEL RULES (PRIOR COHORT)
    # =============================
    # Excel updates:
    # - Loss/RiskAdj/PME/ULAE: IF(AND(Cohort < YEAR(Valuation), Incurred <= Valuation), 0, XLOOKUP(...))
    # - Cancellation: IF(Incurred <= Valuation, 0, 1-(1-annual_cancel)^(1/4))
    # - Premium refund: IF(Incurred <= Valuation, 0, XLOOKUP(...))
    cf['Valuation'] = pd.to_datetime(cf['Valuation'], errors='coerce')
    cf['Incurred'] = pd.to_datetime(cf['Incurred'], errors='coerce')
    cf['Cohort'] = pd.to_numeric(cf['Cohort'], errors='coerce')

    valuation_year = cf['Valuation'].dt.year
    incurred_le_valuation = cf['Incurred'].le(cf['Valuation'])
    prior_cohort_and_past = (cf['Cohort'] < valuation_year) & incurred_le_valuation

    gated_ratio_cols = [
        'Loss_Ratio',
        'Risk_Adjustment_Ratio',
        'PME_Ratio',
        'ULAE_Ratio',
    ]
    for col in gated_ratio_cols:
        if col in cf.columns:
            cf[col] = pd.to_numeric(cf[col], errors='coerce').fillna(0.0)
            cf.loc[prior_cohort_and_past, col] = 0.0

    if 'Premium_Refund_Ratio' in cf.columns:
        cf['Premium_Refund_Ratio'] = pd.to_numeric(cf['Premium_Refund_Ratio'], errors='coerce').fillna(0.0)
        cf.loc[incurred_le_valuation, 'Premium_Refund_Ratio'] = 0.0

    if 'Cancellation_Ratio' in cf.columns:
        annual_cancel = pd.to_numeric(cf['Cancellation_Ratio'], errors='coerce').fillna(0.0)
        # auto-scale percent units (e.g. 5 for 5%)
        mask_pct = annual_cancel.abs() > 1.0
        annual_cancel.loc[mask_pct] = annual_cancel.loc[mask_pct] / 100.0
        annual_cancel = annual_cancel.clip(lower=0.0, upper=1.0)
        quarterly_cancel = 1 - (1 - annual_cancel) ** (1 / 4)
        cf['Cancellation_Ratio'] = 0.0
        cf.loc[~incurred_le_valuation, 'Cancellation_Ratio'] = quarterly_cancel.loc[~incurred_le_valuation]

    # =============================
    # 7. PROBABILITY OF INFORCE (at BoP)
    # =============================
    cf = generate_probability_of_inforce_at_bop(
        cf,
        valuation_col='Valuation',
        incurred_date_col='Incurred',
        cancel_col='Cancellation_Ratio',
        output_col='Probability_of_Inforce_at_BoP',
        icg_col='ICG',
        incurred_seq_col='#Incurred'
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
        prob_inforce_col='Probability_of_Inforce_at_BoP',
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
                prob_inforce_col='Probability_of_Inforce_at_BoP',
				premium_refund_col='Premium_Refund_Ratio',
				cancel_col='Cancellation_Ratio',
				earned_sum_per_icg=earned_sum_per_icg_comm,
				output_col='Exp_Commission',
				earned_premium_col='Earned_Commission',
		)
    

    cf['Exp_Acquisition'] = cf['Expected_Acquisition']

    # =============================
    # 10. INFLATION (Assumptions-based)
    # =============================
    # Excel expects: XLOOKUP(ICG, Assumptions[ICG], Assumptions[Inflation per year (%)])
    inflation_candidates = [
        'Inflation per year (%)',
        'Inflation_per_year (%)',
        'Inflation_per_year',
        'Inflation_Rate',
        'Inflation'
    ]
    annual_infl_col = next((c for c in inflation_candidates if c in assumptions_idx.columns), None)
    if annual_infl_col is None:
        cf['_Annual_Inflation'] = 0.0
    else:
        cf['_Annual_Inflation'] = cf['ICG'].map(assumptions_idx[annual_infl_col])

    cf = generate_inflation_ratio(
        cf,
        annual_inflation_col='_Annual_Inflation',
        valuation_col='Valuation',
        incurred_date_col='Incurred',
        output_col='Inflation_Ratio'
    )

    cf = generate_inflation_factor_at_bop(
        cf,
        inflation_ratio_col='Inflation_Ratio',
        valuation_col='Valuation',
        incurred_date_col='Incurred',
        output_col='Inflation_Factor_at_BoP',
        icg_col='ICG',
        incurred_seq_col='#Incurred'
    )

    if '_Annual_Inflation' in cf.columns:
        cf = cf.drop(columns=['_Annual_Inflation'])

    # =============================
    # 11. EXPENSE + CLAIM (updated formulas)
    # =============================
    cf = generate_expense_claim(
        cf,
        prob_inforce_col='Probability_of_Inforce_at_BoP',
        earned_premium_col='Earned_Premium',
        cancel_col='Cancellation_Ratio',
        ratio_col='PME_Ratio',
        inflation_factor_col='Inflation_Factor_at_BoP',
        inflation_ratio_col='Inflation_Ratio',
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
