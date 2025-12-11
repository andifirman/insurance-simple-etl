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
    # Misalnya data berada di sheet bernama 'InputSheetName'
    current_rate_sheet = all_sheets['Current_Rate']  # Ganti dengan nama sheet yang sesuai
    locked_in_rate_sheet = all_sheets['Locked_in_Rate']
    
    current_rate = current_rate_sheet['Current_Rate']
    locked_in_rate = locked_in_rate_sheet['Locked_in_Rate']
    
    locked_current_rate_data = {
        'Current Rate': current_rate,
        'Locked in Rate': locked_in_rate
    }
    
    return locked_current_rate_data

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
    # Tambahkan Locked in Rate dan Current Rate ke csm_assumption
    csm_assumption['Current Rate'] = locked_current_rate_data['Current Rate']
    csm_assumption['Locked in Rate'] = locked_current_rate_data['Locked in Rate']

		
    csm_expected['Premi SOP'] = generate_premi_sop(csm_expected, all_sheets_dict)
    csm_expected['Premi EOP'] = generate_premi_eop(csm_expected, all_sheets_dict) 
    
    
    
    csm_expected['Expense'] = generate_expense(csm_expected, all_sheets_dict)
    csm_expected['Claim'] = generate_claim(csm_expected, all_sheets_dict)
    csm_expected['Komisi SOP'] = generate_exp_commission_acquisition(csm_expected, all_sheets_dict)
    csm_expected['RA'] = generate_ra(csm_expected, all_sheets_dict)
    csm_expected['Komisi EOP'] = generate_komisi_eop(csm_expected, all_sheets_dict)
    
    csm_actual['Premi'] = generate_actual_premium(csm_actual, all_sheets_dict)
    
     
    csm_actual['Komisi'] = generate_actual_commission_acquisition(csm_actual, all_sheets_dict)


    return {
        "Expected Cashflow": csm_expected,
        "Actual": csm_actual,
        "Assumption": csm_assumption
    }