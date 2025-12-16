# Main application logic placeholder
from src.cashflow.function_date import get_valuation_date
from src.utils.helper_file_picker import get_input_data
from src.cashflow.function_cf_generator import *

from src.etl.etl_extract import load_input_excel
from src.etl.etl_transform import build_cf_gen
from src.etl.etl_load import save_cf_gen

from src.csm.etl_csm_transform import build_csm
from src.csm.etl_csm_load import save_csm_gen

from src.csm.etl_csm_transform import extract_locked_current_rate

def run_app():
  input_file = get_input_data()
  all_sheets = load_input_excel(input_file)
  
  years_forward = int(input("Masukkan jumlah tahun ke depan untuk CF: "))
  
  locked_current_rate_data = extract_locked_current_rate(all_sheets)
  
  # cf = build_cf_gen(all_sheets)
  cf = build_cf_gen(all_sheets, years_forward=years_forward)
  out_path = save_cf_gen(cf, dest_folder="data/processed", output_name="CF_Gen.xlsx")
  
  csm_sheets = build_csm(cf, locked_current_rate_data)
  save_csm_gen(csm_sheets, "data/processed/CSM_Gen.xlsx")