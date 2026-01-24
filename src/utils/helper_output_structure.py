import pandas as pd

def create_cf_gen_template():

    columns = [
        'ICG',
        'Cohort',
        'Portfolio',
        'Contract',
        'Valuation',
        '#Incurred',
        'Incurred',
        'Expected_Premium',
        'Expected_Commission',
        'Expected_Acquisition',
        'Earned_Premium',
        'Earned_Commission',
        'Loss_Ratio',
        'Risk_Adjustment_Ratio',
        'PME_Ratio',
        'ULAE_Ratio',
        'Cancellation_Ratio',
        'Premium_Refund_Ratio',
        'NPR_Ratio',
        'Probability_of_Inforce_at_BoP',
        'Inflation_Ratio',
        'Inflation_Factor_at_BoP',
        'Exp_Premium',
        'Exp_Commission',
        'Exp_Acquisition',
        'Exp_Expense',
        'Exp_Claim',
        'Exp_RA',
        'Exp_NPR',
        'Actual_Premium',
        'Actual_Commission',
        'Actual_Acquisition'
    ]

    return pd.DataFrame(columns=columns)


def create_expected_cashflow_template():
    columns = [
        'Valuation',
        'Incurred',
        'Premi SOP',
        'Premi EOP',
        'Claim',
        'RA', 
        'Expense',
        'Komisi SOP',
        'Komisi EOP',
        'Cohort',
        'LOB',
        'ICG'
    ]

    return pd.DataFrame(columns=columns)


def create_actual_template():
    columns = [
        'Incurred',
        'Premi',
        'Komisi',
        'Cohort',
        'LOB',
        'ICG'
    ]

    return pd.DataFrame(columns=columns)


def create_assumption_template():
    columns = [
        'Valuation',
        'Locked in Rate',
        'Current Rate',
        'Cohort',
        'LOB',
        'ICG'
    ]

    return pd.DataFrame(columns=columns)