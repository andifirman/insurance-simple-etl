# Transform step
import pandas as pd
import numpy as np

from src.csm.function_csm_generator import *
from src.utils.helper_output_structure import *
from src.csm.function_csm_generator import *

# Normalisasi format incurred untuk di CSM
# def normalize_csm_incurred_format(value):
#     try:
#         return pd.to_datetime(value).strftime("%m/%d/%Y")
#     except:
#         return value

def normalize_csm_incurred_format(value):
    try:
        return pd.to_datetime(value, errors='coerce').strftime("%m/%d/%Y") if pd.to_datetime(value, errors='coerce') is not pd.NaT else None
    except Exception as e:
        print(f"Error in date conversion: {e}")
        return None


def extract_locked_current_rate(all_sheets):
    current_rate_sheet = all_sheets.get('Current_Rate')
    locked_in_rate_sheet = all_sheets.get('Locked_in_Rate')

    if current_rate_sheet is None or locked_in_rate_sheet is None:
        raise ValueError("Sheet 'Current_Rate' / 'Locked_in_Rate' tidak ditemukan.")

    for sheet_name, df, needed_cols in [
        ('Current_Rate', current_rate_sheet, {'ICG', 'Current_Rate'}),
        ('Locked_in_Rate', locked_in_rate_sheet, {'ICG', 'Locked_in_Rate'}),
    ]:
        missing = needed_cols - set(df.columns)
        if missing:
            raise ValueError(f"Sheet '{sheet_name}' missing columns: {sorted(missing)}")

    current_map = (
        current_rate_sheet
        .dropna(subset=['ICG'])
        .drop_duplicates(subset=['ICG'])
        .set_index('ICG')['Current_Rate']
    )
    locked_map = (
        locked_in_rate_sheet
        .dropna(subset=['ICG'])
        .drop_duplicates(subset=['ICG'])
        .set_index('ICG')['Locked_in_Rate']
    )

    return {
        'Current Rate': current_map,
        'Locked in Rate': locked_map,
    }

def build_csm(all_sheets_dict, locked_current_rate_data):
    csm_expected = create_expected_cashflow_template()
    csm_actual = create_actual_template()
    csm_assumption = create_assumption_template()
    
    # ============================================= # 
    csm_expected['Valuation'] = all_sheets_dict['Valuation'].apply(normalize_csm_incurred_format)
    csm_expected['Incurred'] = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)

    
    csm_actual['Incurred'] = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    
    csm_assumption['Valuation'] = all_sheets_dict['Valuation'].apply(normalize_csm_incurred_format)
    
    # ============================================= #
    csm_expected['LOB'] = all_sheets_dict['Portfolio']
    csm_actual['LOB'] = all_sheets_dict['Portfolio']
    csm_assumption['LOB'] = all_sheets_dict['Portfolio']
    
    # ============================================= #
    csm_expected['Cohort'] = all_sheets_dict['Cohort']
    csm_actual['Cohort'] = all_sheets_dict['Cohort']
    csm_assumption['Cohort'] = all_sheets_dict['Cohort']
    
    # ============================================= #
    csm_expected['ICG'] = all_sheets_dict['ICG']
    csm_actual['ICG'] = all_sheets_dict['ICG']
    csm_assumption['ICG'] = all_sheets_dict['ICG']

    # ============================================= #
    # Tambahkan Locked in Rate dan Current Rate ke csm_assumption (map by ICG)
    csm_assumption['Current Rate'] = (
        csm_assumption['ICG']
        .map(locked_current_rate_data['Current Rate'])
        .fillna(0)
    )
    csm_assumption['Locked in Rate'] = (
        csm_assumption['ICG']
        .map(locked_current_rate_data['Locked in Rate'])
        .fillna(0)
    )

		
    csm_expected['Premi SOP'] = generate_premi_sop(csm_expected, all_sheets_dict)
    csm_expected['Premi EOP'] = generate_premi_eop(csm_expected, all_sheets_dict) 
    
    
    
    csm_expected['Expense'] = generate_expense(csm_expected, all_sheets_dict)
    csm_expected['Claim'] = generate_claim(csm_expected, all_sheets_dict)
    csm_expected['Komisi SOP'] = generate_exp_commission_acquisition(csm_expected, all_sheets_dict)
    csm_expected['RA'] = generate_ra(csm_expected, all_sheets_dict)
    csm_expected['Komisi EOP'] = generate_komisi_eop(csm_expected, all_sheets_dict)
    
    
    csm_actual['Actual Premi SOP'] = generate_actual_premium(csm_actual, all_sheets_dict)
    csm_actual['Actual Komisi SOP'] = generate_actual_commission_acquisition(csm_actual, all_sheets_dict)

    # Kolom tambahan (belum ada formula dari user) -> default 0 dulu
    for col in [
        'Actual Premi EOP',
        'Actual Komisi EOP',
        'Actual Paid Claim',
        'Actual Paid Expense',
        'Investment Income',
    ]:
        if col not in csm_actual.columns:
            csm_actual[col] = 0
        else:
            csm_actual[col] = 0


    return {
        "Expected Cashflow": csm_expected,
        "Actual": csm_actual,
        "Assumption": csm_assumption
    }