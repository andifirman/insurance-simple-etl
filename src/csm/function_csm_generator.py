import pandas as pd

def normalize_csm_incurred_format(value):
    try:
        dt = pd.to_datetime(value, errors='coerce')
        if pd.isna(dt):
            return None
        # samakan format ke string US mm/dd/yyyy, sama seperti di Excel CSM
        return dt.strftime("%m/%d/%Y")
    except Exception as e:
        print(f"Error in date conversion: {e}")
        return None


def generate_premi_sop(csm_expected, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_expected
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    # Ambil data Exp_Premium dan Incurred dari all_sheets_dict (CF_Gen)
    exp_premium = all_sheets_dict['Exp_Premium']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    # Normalisasi tanggal di kedua sisi ke format yang sama
    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    premi_sop_values = []

    # Excel (jika memang): =-IFERROR(HLOOKUP("Exp_Premium", ...), 0)
    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                premi_sop_values.append(0)
            else:
                idx = match_idx[0]
                val = exp_premium.loc[idx]
                # minus di depan
                premi_sop_values.append(val)

        except Exception:
            premi_sop_values.append(0)

    return premi_sop_values


def generate_claim(csm_expected, all_sheets_dict):
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    expected_claim = all_sheets_dict['Exp_Claim']
    # kalau kamu sudah punya kolom Exp_NPR di CF_Gen, pakai ini:
    # expected_npr = all_sheets_dict['Exp_NPR']
    # tapi sesuai kode kamu sebelumnya, sementara ini:
    expected_npr = all_sheets_dict['Actual_Acquisition']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    results = []

    # Excel: =-Exp_Claim - Exp_NPR
    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                results.append(0)
            else:
                idx = match_idx[0]
                claim_val = expected_claim.loc[idx]
                npr_val = expected_npr.loc[idx]
                results.append(-(claim_val + npr_val))

        except Exception:
            results.append(0)

    return results


def generate_expense(csm_expected, all_sheets_dict):
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    exp_expense = all_sheets_dict['Exp_Expense']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    expense_values = []

    # Excel: =-IFERROR(HLOOKUP("Exp_Expense", ...), 0)
    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                expense_values.append(0)
            else:
                idx = match_idx[0]
                val = exp_expense.loc[idx]
                expense_values.append(-val)

        except Exception:
            expense_values.append(0)

    return expense_values


def generate_ra(csm_expected, all_sheets_dict):
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    exp_ra = all_sheets_dict['Exp_RA']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    ra_values = []

    # Excel: =-IFERROR(HLOOKUP("Exp_RA", ...), 0)
    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                ra_values.append(0)
            else:
                idx = match_idx[0]
                val = exp_ra.loc[idx]
                ra_values.append(-val)

        except Exception:
            ra_values.append(0)

    return ra_values


def generate_actual_premium(csm_actual, all_sheets_dict):
    incurred_values = csm_actual['Incurred']
    icg_values = csm_actual['ICG']

    actual_premium = all_sheets_dict['Actual_Premium']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    premi_values = []

    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                premi_values.append(0)
            else:
                idx = match_idx[0]
                val = actual_premium.loc[idx]
                premi_values.append(val)

        except Exception:
            premi_values.append(0)

    return premi_values


def generate_exp_commission_acquisition(csm_expected, all_sheets_dict):
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    expected_commission = all_sheets_dict['Exp_Commission']
    expected_acquisition = all_sheets_dict['Exp_Acquisition']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    results = []

    # Excel (jika formula-nya): =-Exp_Commission - Exp_Acquisition
    # sekarang kamu punya comm_val - acq_val; kita sesuaikan ke Excel
    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                results.append(0)
            else:
                idx = match_idx[0]
                comm_val = expected_commission.loc[idx]
                acq_val = expected_acquisition.loc[idx]
                results.append(-(comm_val + acq_val))

        except Exception:
            results.append(0)

    return results

# =============================================================== #

def generate_premi_eop(csm_expected, all_sheets_dict):
    # List untuk menyimpan nilai hasil
    premi_eop_values = []
    
    # Ambil data yang diperlukan
    cohort_values = csm_expected['Cohort']
    valuation_values = csm_expected['Valuation']
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']  # Menggunakan ICG untuk agregasi
    
    # Ambil data Actual_Premium dan Incurred dari all_sheets_dict
    actual_premium = all_sheets_dict['Actual_Premium']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)  # Pastikan format yang konsisten
    
    # Loop untuk setiap ICG yang unik
    for icg in csm_expected['ICG'].unique():
        # Filter csm_expected berdasarkan ICG yang sedang diproses
        filtered_df = csm_expected[csm_expected['ICG'] == icg]
        
        # Loop untuk setiap row di csm_expected yang sudah difilter
        for idx in filtered_df.index:
            try:
                # Ambil nilai yang sesuai di baris ini
                cohort = filtered_df.loc[idx, 'Cohort']
                valuation = filtered_df.loc[idx, 'Valuation']
                incurred = filtered_df.loc[idx, 'Incurred']
                
                # Pastikan 'Incurred' diconvert menjadi datetime
                incurred = pd.to_datetime(incurred, errors='coerce')  # Jika gagal, akan menjadi NaT
                
                # Cek kondisi pertama: $J2 < YEAR($A2) dan $B2 <= $A2
                if cohort < pd.to_datetime(valuation).year and incurred <= pd.to_datetime(valuation):
                    # Cari posisi baris yang sesuai dengan nilai Incurred di all_sheets_dict['Incurred']
                    match_idx = incurred_all_sheets[incurred_all_sheets == incurred].index
                    
                    if len(match_idx) > 0:
                        match_idx = match_idx[0]  # Ambil index pertama yang cocok
                        # Ambil nilai Actual_Premium berdasarkan index yang ditemukan
                        actual_premium_value = actual_premium.iloc[match_idx] if match_idx < len(actual_premium) else 0
                        premi_eop_values.append(actual_premium_value)
                    else:
                        premi_eop_values.append(0)
                else:
                    premi_eop_values.append(0)
            except Exception as e:
                print(f"Error encountered at row {idx}: {e}")
                premi_eop_values.append(0)
    
    return premi_eop_values

    
  

def generate_komisi_eop(csm_expected, all_sheets_dict):
    try:
        # Ambil nilai yang diperlukan dari csm_expected
        incurred_values = csm_expected['Incurred']
        cohort_values = csm_expected['Cohort']
        valuation_values = pd.to_datetime(csm_expected['Valuation'], errors='coerce')

        # Data yang akan digunakan dari all_sheets_dict
        actual_commission = all_sheets_dict['Actual_Commission']
        actual_acquisition = all_sheets_dict['Actual_Acquisition']
        incurred_all_sheets = pd.to_datetime(all_sheets_dict['Incurred'], errors='coerce')

        # List untuk menyimpan hasil perhitungan
        commission_and_acquisition_values = []

        # Loop per baris untuk mencocokkan kriteria
        for i, incurred_raw in enumerate(incurred_values):
            try:
                cohort = cohort_values[i]
                valuation = valuation_values[i]

                # Pastikan 'Incurred' di baris CSM berupa datetime
                incurred_value = pd.to_datetime(incurred_raw, errors='coerce')

                # Skip kalau valuation / incurred tidak valid
                if pd.isna(valuation) or pd.isna(incurred_value):
                    commission_and_acquisition_values.append(0)
                    continue

                # Kondisi: $J2 < YEAR($A2) dan $B2 <= $A2
                if cohort < valuation.year and incurred_value <= valuation:
                    # Cari index yang cocok pada kolom 'Incurred'
                    match_idx = incurred_all_sheets[incurred_all_sheets == incurred_value].index

                    if len(match_idx) == 0:
                        commission_and_acquisition_values.append(0)
                    else:
                        idx = match_idx[0]

                        commission_value = actual_commission.loc[idx] if idx in actual_commission.index else 0
                        acquisition_value = actual_acquisition.loc[idx] if idx in actual_acquisition.index else 0

                        # Hasil: - Actual_Commission - Actual_Acquisition
                        result = -(commission_value + acquisition_value)
                        commission_and_acquisition_values.append(result)
                else:
                    commission_and_acquisition_values.append(0)

            except Exception as e:
                print(f"Error encountered at row {i}: {e}")
                commission_and_acquisition_values.append(0)

        return commission_and_acquisition_values

    except Exception as e:
        print(f"General error encountered: {e}")
        return [0] * len(csm_expected)
      
      
def generate_actual_commission_acquisition(csm_actual, all_sheets_dict):
    incurred_values = csm_actual['Incurred']
    icg_values = csm_actual['ICG']

    actual_commission = all_sheets_dict['Actual_Commission']
    actual_acquisition = all_sheets_dict['Actual_Acquisition']
    incurred_all_sheets = all_sheets_dict['Incurred']
    icg_all_sheets = all_sheets_dict['ICG']

    incurred_all_sheets_norm = incurred_all_sheets.apply(normalize_csm_incurred_format)
    incurred_values_norm = incurred_values.apply(normalize_csm_incurred_format)

    results = []

    for incurred_norm, icg_value in zip(incurred_values_norm, icg_values):
        try:
            match_idx = incurred_all_sheets_norm[
                (incurred_all_sheets_norm == incurred_norm) & (icg_all_sheets == icg_value)
            ].index

            if len(match_idx) == 0:
                results.append(0)
            else:
                commission_val = actual_commission.loc[match_idx[0]]
                acquisition_val = actual_acquisition.loc[match_idx[0]]
                results.append(commission_val - acquisition_val)

        except Exception:
            results.append(0)

    return results