# Main application logic placeholder
import pandas as pd

from src.cashflow.function_date import get_valuation_date
from src.utils.helper_file_picker import get_input_data
from src.cashflow.function_cf_generator import *

from src.etl.etl_extract import load_input_excel
from src.etl.etl_transform import build_cf_gen
from src.etl.etl_load import save_cf_gen

from src.csm.etl_csm_transform import build_csm
from src.csm.etl_csm_load import save_csm_gen

from src.csm.etl_csm_transform import extract_locked_current_rate


def _prompt_date(prompt: str, default: pd.Timestamp) -> pd.Timestamp:
  raw = input(prompt).strip()
  if raw == "":
    return pd.to_datetime(default).normalize()
  try:
    return pd.to_datetime(raw, dayfirst=True).normalize()
  except Exception as e:
    raise ValueError(f"Format tanggal tidak valid: '{raw}'. Gunakan yyyy-mm-dd atau dd/mm/yyyy.") from e

def run_app():
  input_file = get_input_data()
  all_sheets = load_input_excel(input_file)

  df_valuation = all_sheets.get('ValuationDate')
  if df_valuation is None:
    raise ValueError("Sheet 'ValuationDate' tidak ditemukan.")
  valuation_date = get_valuation_date(df_valuation)
  
  default_start = pd.Timestamp(year=valuation_date.year, month=1, day=1)
  default_end = pd.Timestamp(year=valuation_date.year, month=12, day=31)

  start_date = _prompt_date(
    f"Masukkan Start Date (default {default_start.date()}, Enter untuk pakai default): ",
    default_start,
  )
  end_date = _prompt_date(
    "Masukkan End Date (format yyyy-mm-dd atau dd/mm/yyyy, Enter untuk pakai default): ",
    default_end,
  )

  if start_date.year > end_date.year:
    raise ValueError(f"Start Date ({start_date.date()}) tidak boleh setelah End Date ({end_date.date()}).")
  
  claim_inflation_percent = float(input("Masukkan inflasi klaim (mis. 5 untuk 5%): "))
  expense_inflation_percent = float(input("Masukkan inflasi expense (mis. 5 untuk 5%): "))
  
  claim_inflation_rate = claim_inflation_percent / 100.0
  expense_inflation_rate = expense_inflation_percent / 100.0
  
  
  locked_current_rate_data = extract_locked_current_rate(all_sheets)
  
  # cf = build_cf_gen(all_sheets)
  cf = build_cf_gen(
    all_sheets, 
    start_date=start_date,
    end_date=end_date,
    claim_inflation_rate=claim_inflation_rate, 
    expense_inflation_rate=expense_inflation_rate
  )
  
  out_path = save_cf_gen(cf, dest_folder="data/processed", output_name="CF_Gen.xlsx")
  
  csm_sheets = build_csm(cf, locked_current_rate_data)
  save_csm_gen(csm_sheets, "data/processed/CSM_Gen.xlsx")