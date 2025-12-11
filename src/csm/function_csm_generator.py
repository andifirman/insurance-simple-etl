import pandas as pd

def normalize_csm_incurred_format(value):
    try:
        return pd.to_datetime(value, errors='coerce').strftime("%m/%d/%Y") if pd.to_datetime(value, errors='coerce') is not pd.NaT else None
    except Exception as e:
        print(f"Error in date conversion: {e}")
        return None


def generate_premi_sop(csm_expected, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_expected
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    # Ambil data Exp_Premium dan Incurred dari all_sheets_dict
    exp_premium = all_sheets_dict['Exp_Premium']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil pencocokan
    premi_sop_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_expected
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index
            
            if len(match_idx) == 0:
                premi_sop_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Exp_Premium berdasarkan index yang ditemukan
                exp_premium_value = exp_premium.loc[match_idx[0]]
                premi_sop_values.append(exp_premium_value)

        except IndexError:
            # Jika tidak ditemukan, kita bisa mengembalikan 0 atau nilai lain sesuai kebutuhan
            premi_sop_values.append(0)

    return premi_sop_values

  

def generate_expense(csm_expected, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_expected
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    # Ambil data Exp_Expense dan Incurred dari all_sheets_dict
    exp_expense = all_sheets_dict['Exp_Expense']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil pencocokan
    expense_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_expected
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index
            
            if len(match_idx) == 0:
                expense_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Exp_Expense berdasarkan index yang ditemukan
                exp_expense_value = exp_expense.loc[match_idx[0]]  # Ambil nilai Exp_Expense yang sesuai
                expense_values.append(exp_expense_value)

        except IndexError:
            # Jika tidak ditemukan, kita bisa mengembalikan 0 atau nilai lain sesuai kebutuhan
            expense_values.append(0)

    return expense_values

def generate_ra(csm_expected, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_expected
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    # Ambil data Exp_Expense dan Incurred dari all_sheets_dict
    exp_ra = all_sheets_dict['Exp_RA']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil pencocokan
    ra_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_expected
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index
            
            if len(match_idx) == 0:
                ra_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Exp_Expense berdasarkan index yang ditemukan
                exp_ra_value = exp_ra.loc[match_idx[0]]  # Ambil nilai Exp_Expense yang sesuai
                ra_values.append(exp_ra_value)

        except IndexError:
            # Jika tidak ditemukan, kita bisa mengembalikan 0 atau nilai lain sesuai kebutuhan
            ra_values.append(0)

    return ra_values

  
  
def generate_actual_premium(csm_actual, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_actual
    incurred_values = csm_actual['Incurred']
    icg_values = csm_actual['ICG']

    # Ambil data Actual_Premium dan Incurred dari all_sheets_dict
    exp_premium = all_sheets_dict['Actual_Premium']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil pencocokan
    premi_sop_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_actual
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index

            if len(match_idx) == 0:
                premi_sop_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Actual_Premium berdasarkan index yang ditemukan
                exp_premium_value = exp_premium.loc[match_idx[0]]  # Ambil nilai Actual_Premium yang sesuai
                premi_sop_values.append(exp_premium_value)

        except IndexError:
            # Jika tidak ditemukan, kita bisa mengembalikan 0 atau nilai lain sesuai kebutuhan
            premi_sop_values.append(0)

    return premi_sop_values


def generate_actual_commission_acquisition(csm_actual, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_actual
    incurred_values = csm_actual['Incurred']
    icg_values = csm_actual['ICG']

    # Ambil data Actual_Commission dan Actual_Acquisition dari all_sheets_dict
    actual_commission = all_sheets_dict['Actual_Commission']
    actual_acquisition = all_sheets_dict['Actual_Acquisition']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil perhitungan
    commission_acquisition_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_actual
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index
            
            if len(match_idx) == 0:
                commission_acquisition_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Actual_Commission dan Actual_Acquisition berdasarkan index yang ditemukan
                actual_commission_value = actual_commission.loc[match_idx[0]]
                actual_acquisition_value = actual_acquisition.loc[match_idx[0]]

                # Hitung hasil perhitungan (Actual_Commission - Actual_Acquisition)
                result_value = actual_commission_value - actual_acquisition_value
                commission_acquisition_values.append(result_value)

        except IndexError:
            commission_acquisition_values.append(0)  # Jika terjadi error, kita append 0

    return commission_acquisition_values


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
  
  
def generate_claim(csm_expected, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_actual
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    # Ambil data Actual_Commission dan Actual_Acquisition dari all_sheets_dict
    expected_claim = all_sheets_dict['Exp_Claim']
    expected_npr = all_sheets_dict['Actual_Acquisition']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil perhitungan
    commission_acquisition_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_actual
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index
            
            if len(match_idx) == 0:
                commission_acquisition_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Actual_Commission dan Actual_Acquisition berdasarkan index yang ditemukan
                expected_claim_value = expected_claim.loc[match_idx[0]]
                expected_acquisition_value = expected_npr.loc[match_idx[0]]

                # Hitung hasil perhitungan (Actual_Commission - Actual_Acquisition)
                result_value = expected_claim_value - expected_acquisition_value
                commission_acquisition_values.append(result_value)

        except IndexError:
            commission_acquisition_values.append(0)  # Jika terjadi error, kita append 0

    return commission_acquisition_values
  
  
def generate_exp_commission_acquisition(csm_expected, all_sheets_dict):
    # Ambil nilai Incurred dan ICG dari csm_actual
    incurred_values = csm_expected['Incurred']
    icg_values = csm_expected['ICG']

    # Ambil data Actual_Commission dan Actual_Acquisition dari all_sheets_dict
    expected_commission = all_sheets_dict['Exp_Commission']
    expected_acquisition = all_sheets_dict['Exp_Acquisition']
    incurred_all_sheets = all_sheets_dict['Incurred'].apply(normalize_csm_incurred_format)
    icg_all_sheets = all_sheets_dict['ICG']

    # List untuk menyimpan nilai hasil perhitungan
    commission_acquisition_values = []

    # Loop untuk setiap nilai Incurred dan ICG di csm_actual
    for incurred_value, icg_value in zip(incurred_values, icg_values):
        try:
            # Cari posisi baris yang sesuai dengan Incurred dan ICG
            match_idx = incurred_all_sheets[(incurred_all_sheets == incurred_value) & (icg_all_sheets == icg_value)].index
            
            if len(match_idx) == 0:
                commission_acquisition_values.append(0)  # Jika tidak ditemukan, kita append 0
            else:
                # Ambil nilai Actual_Commission dan Actual_Acquisition berdasarkan index yang ditemukan
                expected_commission_value = expected_commission.loc[match_idx[0]]
                expected_acquisition_value = expected_acquisition.loc[match_idx[0]]

                # Hitung hasil perhitungan (Actual_Commission - Actual_Acquisition)
                result_value = expected_commission_value - expected_acquisition_value
                commission_acquisition_values.append(result_value)

        except IndexError:
            commission_acquisition_values.append(0)  # Jika terjadi error, kita append 0

    return commission_acquisition_values
  

def generate_komisi_eop(csm_expected, all_sheets_dict):
    try:
        # Ambil nilai yang diperlukan dari csm_expected
        incurred_values = csm_expected['Incurred']
        cohort_values = csm_expected['Cohort']
        valuation_values = csm_expected['Valuation']
        
        # Data yang akan digunakan dari all_sheets_dict
        actual_commission = all_sheets_dict['Actual_Commission']
        actual_acquisition = all_sheets_dict['Actual_Acquisition']
        incurred_all_sheets = pd.to_datetime(all_sheets_dict['Incurred'], errors='coerce')

        # Pastikan 'Valuation' di csm_expected juga diubah ke format datetime yang seragam
        valuation_values = pd.to_datetime(csm_expected['Valuation'], errors='coerce')

        # List untuk menyimpan hasil perhitungan
        commission_and_acquisition_values = []

        # Loop per baris untuk mencocokkan kriteria
        for i, incurred_value in enumerate(incurred_values):
            try:
                # Ambil Cohort dan Valuation dari baris yang sama
                cohort = cohort_values[i]
                valuation = valuation_values[i]
                
                # Pastikan 'Incurred' juga dalam format datetime yang sesuai
                incurred_value = pd.to_datetime(incurred_value, errors='coerce')

                # Periksa kondisi $J2<YEAR($A2),$B2<=$A2
                if cohort < valuation.year and incurred_value <= valuation:
                    # Cari index yang cocok pada kolom 'Incurred' untuk mencari nilai Actual_Commission dan Actual_Acquisition
                    idx = incurred_all_sheets[incurred_all_sheets == incurred_value].index[0]
                    
                    # Ambil nilai Actual_Commission dan Actual_Acquisition dari index yang ditemukan
                    commission_value = actual_commission.loc[idx] if idx < len(actual_commission) else 0
                    acquisition_value = actual_acquisition.loc[idx] if idx < len(actual_acquisition) else 0
                    
                    # Kembalikan hasil perhitungan: - Actual_Commission - Actual_Acquisition
                    result = -(commission_value + acquisition_value)
                    commission_and_acquisition_values.append(result)
                else:
                    # Jika kondisi tidak terpenuhi, kembalikan 0
                    commission_and_acquisition_values.append(0)
            except Exception as e:
                # Tangani error jika ada masalah dengan pencarian index atau perhitungan
                print(f"Error encountered at row {i}: {e}")
                commission_and_acquisition_values.append(0)

        return commission_and_acquisition_values
    
    except Exception as e:
        print(f"General error encountered: {e}")
        return [0] * len(csm_expected)  # Kembalikan list berisi 0 jika error
