import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.etl.etl_load import save_cf_gen
from src.csm.etl_csm_load import save_csm_gen


def build_example_cf_gen() -> pd.DataFrame:
    icg = "ICG-AAA"
    cohort = 2022
    valuation = pd.Timestamp("2023-12-31")

    incurred_dates = pd.to_datetime(
        ["2023-03-31", "2023-06-30", "2023-09-30", "2023-12-31"]
    )

    assumption_set_id = ["A", "A", "B", "B"]
    assumption_effective_from = pd.to_datetime(
        ["2023-01-01", "2023-01-01", "2023-07-01", "2023-07-01"]
    )

    df = pd.DataFrame(
        {
            "ICG": [icg] * 4,
            "Cohort": [cohort] * 4,
            "Valuation": [valuation] * 4,
            "Incurred": incurred_dates,
            "#Incurred": [1, 2, 3, 4],
            "AssumptionSetId": assumption_set_id,
            "Assumption_Effective_From": assumption_effective_from,
            "Current_Rate": [0.0666, 0.0666, 0.0654, 0.0654],
            "Expense_Ratio": [0.10, 0.10, 0.12, 0.12],
            # contoh kolom-kolom yang biasanya ada di CF Gen
            "Expected_Premium": [1000, 1000, 1000, 1000],
            "Expected_Commission": [-30, -30, -30, -30],
            "Expected_Acquisition": [-10, -10, -10, -10],
            "Earned_Premium": [1000, 1000, 1000, 1000],
            "Earned_Commission": [-30, -30, -30, -30],
            "Exp_Premium": [950, 900, 850, 800],
            "Exp_Commission": [-28, -27, -26, -25],
            "Exp_Acquisition": [-9, -9, -9, -9],
            "Exp_Expense": [-95, -90, -102, -96],
            "Exp_Claim": [-300, -280, -260, -240],
            "Exp_RA": [-20, -18, -17, -15],
            "Exp_NPR": [-5, -5, -5, -5],
            "Actual_Premium": [0, 0, 0, 0],
            "Actual_Commission": [0, 0, 0, 0],
            "Actual_Acquisition": [0, 0, 0, 0],
            "Exp_Claim_Inflation": [-300, -280, -260, -240],
            "Exp_Expense_Inflation": [-95, -90, -102, -96],
            "Probability_of_Inforce": [1.0, 0.98, 0.96, 0.94],
            "NPR_Ratio": [0.005, 0.005, 0.005, 0.005],
        }
    )

    return df


def build_example_csm_dict() -> dict[str, pd.DataFrame]:
    icg = "ICG-AAA"
    cohort = 2022
    valuation = "31-Dec-2023"

    incurred = ["31-Mar-2023", "30-Jun-2023", "30-Sep-2023", "31-Dec-2023"]
    assumption_set_id = ["A", "A", "B", "B"]
    assumption_effective_from = ["01-Jan-2023", "01-Jan-2023", "01-Jul-2023", "01-Jul-2023"]

    expected = pd.DataFrame(
        {
            "ICG": [icg] * 4,
            "Cohort": [cohort] * 4,
            "Valuation": [valuation] * 4,
            "Incurred": incurred,
            "AssumptionSetId": assumption_set_id,
            "Assumption_Effective_From": assumption_effective_from,
            "Premi SOP": [950, 900, 850, 800],
            "Premi EOP": [535, 512, 471, 449],
            "Claim": [-300, -280, -260, -240],
            "RA": [-20, -18, -17, -15],
            "Expense": [-95, -90, -102, -96],
            "Komisi SOP": [-28, -27, -26, -25],
            "Komisi EOP": [-30, -28, -27, -25],
        }
    )

    actual = pd.DataFrame(
        {
            "ICG": [icg] * 4,
            "Cohort": [cohort] * 4,
            "Valuation": [valuation] * 4,
            "Incurred": incurred,
            "Premi": [0, 0, 0, 0],
            "Komisi": [0, 0, 0, 0],
        }
    )

    return {"Expected Cashflow": expected, "Actual": actual}


def main() -> None:
    df_cf = build_example_cf_gen()
    csm_dict = build_example_csm_dict()

    cf_path = save_cf_gen(
        df_cf,
        dest_folder="data/processed",
        output_name="EXAMPLE_mid_assumption_CF_Gen.xlsx",
    )

    csm_path = save_csm_gen(
        csm_dict,
        output_path="data/processed/EXAMPLE_mid_assumption_CSM_Gen.xlsx",
    )

    print(f"Wrote: {cf_path}")
    print(f"Wrote: {csm_path}")


if __name__ == "__main__":
    main()
