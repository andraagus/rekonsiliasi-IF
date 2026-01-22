import pandas as pd
import numpy as np
import os

def rekon_acc():


    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    os_acc = pd.read_excel(os.path.join(base_dir, "OS_ACC.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_acc = pd.read_excel(os.path.join(base_dir, "BPKB_ACC.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})

    # Load data from text files with delimiter '|'
    inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
    write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)



    # # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    # os_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/CORACCOUNT_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # os_acc = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OS_ACC.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OBJECTCAR_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_acc = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/BPKB_ACC.xlsx", dtype={'PARTNER_AGRMNT_NO': str})

    # # Load data from text files with delimiter '|'
    # inventory_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/ASSETINVENTORY.txt", delimiter="|", low_memory=False)
    # write_off_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/WRITEOFF.txt", delimiter="|", low_memory=False)


    # Cek Nama Kolom yang ada di tabel
    print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(os_acc.columns) # Periksa nama kolom yang akan digunakan

    # Filter OS BJJ ACC
    oss_bjj_acc = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('ASF') | os_bjj['PROD_OFFERING_CODE'].str.startswith('SBSF')]

    # Konversi type data agar sama dengan coreaccount
    oss_bjj_acc.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_acc['PARTNER_AGRMNT_NO'].astype(str)
    os_acc['PARTNER_AGRMNT_NO'] = os_acc['PARTNER_AGRMNT_NO'].astype(str)

    recon_JF_ACC = pd.merge(oss_bjj_acc[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_acc[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_ACC']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DIFF_OS
    recon_JF_ACC['DIFF_OS'] = recon_JF_ACC['OS_PRINCIPAL_AMT'] - recon_JF_ACC['OS_PRINCIPAL_AMT_ACC']

    # Menambahkan kolom DATA_STATUS
    conditions = [
        (recon_JF_ACC['_merge'] == 'both'),
        (recon_JF_ACC['_merge'] == 'right_only'),
        (recon_JF_ACC['_merge'] == 'left_only') & (recon_JF_ACC['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
        (recon_JF_ACC['_merge'] == 'left_only') & (recon_JF_ACC['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
        (recon_JF_ACC['_merge'] == 'left_only') & (recon_JF_ACC['CONTRACT_STATUS'] == 'EXP')
    ]
    choices = [
        'Data Tersedia di BJJ & Partner',
        'Data hanya Tersedia di ACC',
        'Data hanya tersedia di BJJ dan Statusnya Inventory',
        'Data hanya tersedia di BJJ dan Statusnya Writeoff',
        'Data hanya tersedia di BJJ dan Statusnya Lunas'
    ]
    recon_JF_ACC['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')

    # Menambahkan kolom DIFF_OS_STATUS
    conditions = [
        (recon_JF_ACC['DIFF_OS'].isna()),
        (recon_JF_ACC['DIFF_OS'] > 100),
        (recon_JF_ACC['DIFF_OS'] < -100),
        (recon_JF_ACC['DIFF_OS'] > 0) & (recon_JF_ACC['DIFF_OS'] <= 100),
        (recon_JF_ACC['DIFF_OS'] < 0) & (recon_JF_ACC['DIFF_OS'] >= -100),
        (recon_JF_ACC['DIFF_OS'] == 0)
    ]
    choices = [
        recon_JF_ACC['DATA_STATUS'],
        'Selisih > 100',
        'Selisih < -100',
        'Selisih +- 100',
        'Selisih +- 100',
        'Selisih 0'
   
    ]
    recon_JF_ACC['DIFF_OS_STATUS'] = np.select(conditions, choices)

    print(recon_JF_ACC)

    # # Membuat pivot table Status availibility data
    pivot_table_data = pd.pivot_table(
        recon_JF_ACC,
        values=['PARTNER_AGRMNT_NO'],
        index=['DATA_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data)



    # # Membuat pivot table selisih OS dan OS Principal
    pivot_table_diff_os = pd.pivot_table(
        recon_JF_ACC,
        values=['DIFF_OS', 'OS_PRINCIPAL_AMT_ACC', 'PARTNER_AGRMNT_NO'],
        index=['DIFF_OS_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_diff_os)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_OS_ACC.xlsx") as writer:
        recon_JF_ACC.to_excel(writer, sheet_name='Hasil Recon', index=False)
        pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
        pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
    print("File Excel berhasil dibuat dengan tiga sheet.")


def rekon_fif():

    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    os_fif = pd.read_excel(os.path.join(base_dir, "OS_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_fif = pd.read_excel(os.path.join(base_dir, "BPKB_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})

    # Load data from text files with delimiter '|'
    inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
    write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)


    # # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    # os_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/CORACCOUNT_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # os_fif = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OS_FIF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OBJECTCAR_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_fif = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/BPKB_FIF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})

    # # Load data from text files with delimiter '|'
    # inventory_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/ASSETINVENTORY.txt", delimiter="|", low_memory=False)
    # write_off_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/WRITEOFF.txt", delimiter="|", low_memory=False)
    os_fif['OS_PRINCIPAL_AMT_FIF'] = pd.to_numeric(os_fif['OS_PRINCIPAL_AMT_FIF'], errors='coerce')

    # Cek Nama Kolom yang ada di tabel
    print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(os_fif.columns) # Periksa nama kolom yang akan digunakan

    # Filter OS BJJ ACC
    oss_bjj_fif = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('FIF') | os_bjj['PROD_OFFERING_CODE'].str.startswith('FIF2')]

    # Konversi type data agar sama dengan coreaccount
    oss_bjj_fif.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_fif['PARTNER_AGRMNT_NO'].astype(str)
    os_fif['PARTNER_AGRMNT_NO'] = os_fif['PARTNER_AGRMNT_NO'].astype(str)

    recon_JF_FIF = pd.merge(oss_bjj_fif[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_fif[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_FIF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DIFF_OS
    recon_JF_FIF['DIFF_OS'] = recon_JF_FIF['OS_PRINCIPAL_AMT'] - recon_JF_FIF['OS_PRINCIPAL_AMT_FIF']

    # Menambahkan kolom DATA_STATUS
    conditions = [
        (recon_JF_FIF['_merge'] == 'both'),
        (recon_JF_FIF['_merge'] == 'right_only'),
        (recon_JF_FIF['_merge'] == 'left_only') & (recon_JF_FIF['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
        (recon_JF_FIF['_merge'] == 'left_only') & (recon_JF_FIF['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
        (recon_JF_FIF['_merge'] == 'left_only') & (recon_JF_FIF['CONTRACT_STATUS'] == 'EXP')
    ]
    choices = [
        'Data Tersedia di BJJ & Partner',
        'Data hanya Tersedia di FIF',
        'Data hanya tersedia di BJJ dan Statusnya Inventory',
        'Data hanya tersedia di BJJ dan Statusnya Writeoff',
        'Data hanya tersedia di BJJ dan Statusnya Lunas'
    ]
    recon_JF_FIF['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')

    # Menambahkan kolom DIFF_OS_STATUS
    conditions = [
        (recon_JF_FIF['DIFF_OS'].isna()),
        (recon_JF_FIF['DIFF_OS'] > 100),
        (recon_JF_FIF['DIFF_OS'] < -100),
        (recon_JF_FIF['DIFF_OS'] > 0) & (recon_JF_FIF['DIFF_OS'] <= 100),
        (recon_JF_FIF['DIFF_OS'] < 0) & (recon_JF_FIF['DIFF_OS'] >= -100),
        (recon_JF_FIF['DIFF_OS'] == 0)
    ]
    choices = [
        recon_JF_FIF['DATA_STATUS'],
        'Selisih > 100',
        'Selisih < -100',
        'Selisih +- 100',
        'Selisih +- 100',
        'Selisih 0'
   
    ]
    recon_JF_FIF['DIFF_OS_STATUS'] = np.select(conditions, choices)

    print(recon_JF_FIF)

    # # Membuat pivot table Status availibility data
    pivot_table_data = pd.pivot_table(
        recon_JF_FIF,
        values=['PARTNER_AGRMNT_NO'],
        index=['DATA_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data)



    # # Membuat pivot table selisih OS dan OS Principal
    pivot_table_diff_os = pd.pivot_table(
        recon_JF_FIF,
        values=['DIFF_OS', 'OS_PRINCIPAL_AMT_FIF', 'PARTNER_AGRMNT_NO'],
        index=['DIFF_OS_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_diff_os)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_OS_FIF.xlsx") as writer:
        recon_JF_FIF.to_excel(writer, sheet_name='Hasil Recon', index=False)
        pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
        pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
    print("File Excel berhasil dibuat dengan tiga sheet.")


def rekon_taf():
 
    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    os_taf = pd.read_excel(os.path.join(base_dir, "OS_TAF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_taf = pd.read_excel(os.path.join(base_dir, "BPKB_TAF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})

    # Load data from text files with delimiter '|'
    inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
    write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)

    # # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    # os_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/CORACCOUNT_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # os_taf = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OS_TAF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OBJECTCAR_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_taf = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/BPKB_TAF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})

    # # Load data from text files with delimiter '|'
    # inventory_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/ASSETINVENTORY.txt", delimiter="|", low_memory=False)
    # write_off_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/WRITEOFF.txt", delimiter="|", low_memory=False)


    # Cek Nama Kolom yang ada di tabel
    print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(os_taf.columns) # Periksa nama kolom yang akan digunakan

    # Filter OS BJJ ACC
    oss_bjj_taf = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('TAF') | os_bjj['PROD_OFFERING_CODE'].str.startswith('TAF2')]

    # Konversi type data agar sama dengan coreaccount
    oss_bjj_taf.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_taf['PARTNER_AGRMNT_NO'].astype(str)
    os_taf['PARTNER_AGRMNT_NO'] = os_taf['PARTNER_AGRMNT_NO'].astype(str)

    recon_JF_TAF = pd.merge(oss_bjj_taf[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_taf[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_TAF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DIFF_OS
    recon_JF_TAF['DIFF_OS'] = recon_JF_TAF['OS_PRINCIPAL_AMT'] - recon_JF_TAF['OS_PRINCIPAL_AMT_TAF']

    # Menambahkan kolom DATA_STATUS
    conditions = [
        (recon_JF_TAF['_merge'] == 'both'),
        (recon_JF_TAF['_merge'] == 'right_only'),
        (recon_JF_TAF['_merge'] == 'left_only') & (recon_JF_TAF['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
        (recon_JF_TAF['_merge'] == 'left_only') & (recon_JF_TAF['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
        (recon_JF_TAF['_merge'] == 'left_only') & (recon_JF_TAF['CONTRACT_STATUS'] == 'EXP')
    ]
    choices = [
        'Data Tersedia di BJJ & Partner',
        'Data hanya Tersedia di TAF',
        'Data hanya tersedia di BJJ dan Statusnya Inventory',
        'Data hanya tersedia di BJJ dan Statusnya Writeoff',
        'Data hanya tersedia di BJJ dan Statusnya Lunas'
    ]
    recon_JF_TAF['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')

    # Menambahkan kolom DIFF_OS_STATUS
    conditions = [
        (recon_JF_TAF['DIFF_OS'].isna()),
        (recon_JF_TAF['DIFF_OS'] > 100),
        (recon_JF_TAF['DIFF_OS'] < -100),
        (recon_JF_TAF['DIFF_OS'] > 0) & (recon_JF_TAF['DIFF_OS'] <= 100),
        (recon_JF_TAF['DIFF_OS'] < 0) & (recon_JF_TAF['DIFF_OS'] >= -100),
        (recon_JF_TAF['DIFF_OS'] == 0)
    ]
    choices = [
        recon_JF_TAF['DATA_STATUS'],
        'Selisih > 100',
        'Selisih < -100',
        'Selisih +- 100',
        'Selisih +- 100',
        'Selisih 0'
    ]
    recon_JF_TAF['DIFF_OS_STATUS'] = np.select(conditions, choices)

    print(recon_JF_TAF)

    # # Membuat pivot table Status availibility data
    pivot_table_data = pd.pivot_table(
        recon_JF_TAF,
        values=['PARTNER_AGRMNT_NO'],
        index=['DATA_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data)

    # # Membuat pivot table selisih OS dan OS Principal
    pivot_table_diff_os = pd.pivot_table(
        recon_JF_TAF,
        values=['DIFF_OS', 'OS_PRINCIPAL_AMT_TAF', 'PARTNER_AGRMNT_NO'],
        index=['DIFF_OS_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_diff_os)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_OS_TAF.xlsx") as writer:
        recon_JF_TAF.to_excel(writer, sheet_name='Hasil Recon', index=False)
        pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
        pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
    print("File Excel berhasil dibuat dengan tiga sheet.")

def rekon_amartha():
 
    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    os_amartha = pd.read_excel(os.path.join(base_dir, "OS_AMARTHA.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})

    # Load data from text files with delimiter '|'
    inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
    write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)

    # Cek Nama Kolom yang ada di tabel
    print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(os_amartha.columns) # Periksa nama kolom yang akan digunakan

    # Filter OS BJJ ACC
    oss_bjj_amartha = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('AMARTHA') | os_bjj['PROD_OFFERING_CODE'].str.startswith('AMARTHA')]

    # Konversi type data agar sama dengan coreaccount
    oss_bjj_amartha.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_amartha['PARTNER_AGRMNT_NO'].astype(str)
    os_amartha['PARTNER_AGRMNT_NO'] = os_amartha['PARTNER_AGRMNT_NO'].astype(str)

    recon_JF_AMARTHA = pd.merge(oss_bjj_amartha[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_amartha[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_TAF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DIFF_OS
    recon_JF_AMARTHA['DIFF_OS'] = recon_JF_AMARTHA['OS_PRINCIPAL_AMT'] - recon_JF_AMARTHA['OS_PRINCIPAL_AMT_AMARTHA']

    # Menambahkan kolom DATA_STATUS
    conditions = [
        (recon_JF_AMARTHA['_merge'] == 'both'),
        (recon_JF_AMARTHA['_merge'] == 'right_only'),
        (recon_JF_AMARTHA['_merge'] == 'left_only') & (recon_JF_AMARTHA['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
        (recon_JF_AMARTHA['_merge'] == 'left_only') & (recon_JF_AMARTHA['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
        (recon_JF_AMARTHA['_merge'] == 'left_only') & (recon_JF_AMARTHA['CONTRACT_STATUS'] == 'EXP')
    ]
    choices = [
        'Data Tersedia di BJJ & Partner',
        'Data hanya Tersedia di AMARTHA',
        'Data hanya tersedia di BJJ dan Statusnya Inventory',
        'Data hanya tersedia di BJJ dan Statusnya Writeoff',
        'Data hanya tersedia di BJJ dan Statusnya Lunas'
    ]
    recon_JF_AMARTHA['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')

    # Menambahkan kolom DIFF_OS_STATUS
    conditions = [
        (recon_JF_AMARTHA['DIFF_OS'].isna()),
        (recon_JF_AMARTHA['DIFF_OS'] > 100),
        (recon_JF_AMARTHA['DIFF_OS'] < -100),
        (recon_JF_AMARTHA['DIFF_OS'] > 0) & (recon_JF_AMARTHA['DIFF_OS'] <= 100),
        (recon_JF_AMARTHA['DIFF_OS'] < 0) & (recon_JF_AMARTHA['DIFF_OS'] >= -100),
        (recon_JF_AMARTHA['DIFF_OS'] == 0)
    ]
    choices = [
        recon_JF_AMARTHA['DATA_STATUS'],
        'Selisih > 100',
        'Selisih < -100',
        'Selisih +- 100',
        'Selisih +- 100',
        'Selisih 0'
   
    ]
    recon_JF_AMARTHA['DIFF_OS_STATUS'] = np.select(conditions, choices)

    print(recon_JF_AMARTHA)

    # # Membuat pivot table Status availibility data
    pivot_table_data = pd.pivot_table(
        recon_JF_AMARTHA,
        values=['PARTNER_AGRMNT_NO'],
        index=['DATA_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data)

    # # Membuat pivot table selisih OS dan OS Principal
    pivot_table_diff_os = pd.pivot_table(
        recon_JF_AMARTHA,
        values=['DIFF_OS', 'OS_PRINCIPAL_AMT_TAF', 'PARTNER_AGRMNT_NO'],
        index=['DIFF_OS_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_diff_os)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_OS_AMARTHA.xlsx") as writer:
        recon_JF_AMARTHA.to_excel(writer, sheet_name='Hasil Recon', index=False)
        pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
        pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
    print("File Excel berhasil dibuat dengan tiga sheet.")

def rekon_easycash():
 
    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    os_easycash = pd.read_excel(os.path.join(base_dir, "OS_EASYCASH.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})

    # Load data from text files with delimiter '|'
    inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
    write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)

    # Cek Nama Kolom yang ada di tabel
    print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(os_easycash.columns) # Periksa nama kolom yang akan digunakan

    # Filter OS BJJ ACC
    oss_bjj_easycash = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('EASYCASH') | os_bjj['PROD_OFFERING_CODE'].str.startswith('EASYCASH')]

    # Konversi type data agar sama dengan coreaccount
    oss_bjj_easycash.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_easycash['PARTNER_AGRMNT_NO'].astype(str)
    os_easycash['PARTNER_AGRMNT_NO'] = os_easycash['PARTNER_AGRMNT_NO'].astype(str)

    recon_JF_EASYCASH = pd.merge(oss_bjj_easycash[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_easycash[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_EASYCASH']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DIFF_OS
    recon_JF_EASYCASH['DIFF_OS'] = recon_JF_EASYCASH['OS_PRINCIPAL_AMT'] - recon_JF_EASYCASH['OS_PRINCIPAL_AMT_EASYCASH']

    # Menambahkan kolom DATA_STATUS
    conditions = [
        (recon_JF_EASYCASH['_merge'] == 'both'),
        (recon_JF_EASYCASH['_merge'] == 'right_only'),
        (recon_JF_EASYCASH['_merge'] == 'left_only') & (recon_JF_EASYCASH['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
        (recon_JF_EASYCASH['_merge'] == 'left_only') & (recon_JF_EASYCASH['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
        (recon_JF_EASYCASH['_merge'] == 'left_only') & (recon_JF_EASYCASH['CONTRACT_STATUS'] == 'EXP')
    ]
    choices = [
        'Data Tersedia di BJJ & Partner',
        'Data hanya Tersedia di EASYCASH',
        'Data hanya tersedia di BJJ dan Statusnya Inventory',
        'Data hanya tersedia di BJJ dan Statusnya Writeoff',
        'Data hanya tersedia di BJJ dan Statusnya Lunas'
    ]
    recon_JF_EASYCASH['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')

    # Menambahkan kolom DIFF_OS_STATUS
    conditions = [
        (recon_JF_EASYCASH['DIFF_OS'].isna()),
        (recon_JF_EASYCASH['DIFF_OS'] > 100),
        (recon_JF_EASYCASH['DIFF_OS'] < -100),
        (recon_JF_EASYCASH['DIFF_OS'] > 0) & (recon_JF_EASYCASH['DIFF_OS'] <= 100),
        (recon_JF_EASYCASH['DIFF_OS'] < 0) & (recon_JF_EASYCASH['DIFF_OS'] >= -100),
        (recon_JF_EASYCASH['DIFF_OS'] == 0)
    ]
    choices = [
        recon_JF_EASYCASH['DATA_STATUS'],
        'Selisih > 100',
        'Selisih < -100',
        'Selisih +- 100',
        'Selisih +- 100',
        'Selisih 0'
    ]
    recon_JF_EASYCASH['DIFF_OS_STATUS'] = np.select(conditions, choices)

    print(recon_JF_EASYCASH)

    # # Membuat pivot table Status availibility data
    pivot_table_data = pd.pivot_table(
        recon_JF_EASYCASH,
        values=['PARTNER_AGRMNT_NO'],
        index=['DATA_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data)



    # # Membuat pivot table selisih OS dan OS Principal
    pivot_table_diff_os = pd.pivot_table(
        recon_JF_EASYCASH,
        values=['DIFF_OS', 'OS_PRINCIPAL_AMT_EASYCASH', 'PARTNER_AGRMNT_NO'],
        index=['DIFF_OS_STATUS'],
        columns=['PARTNER_NAME'],
        aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_diff_os)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_OS_EASYCASH.xlsx") as writer:
        recon_JF_EASYCASH.to_excel(writer, sheet_name='Hasil Recon', index=False)
        pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
        pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
    print("File Excel berhasil dibuat dengan tiga sheet.")

# Panggil fungsi rekon_acc, rekon_fif, dan rekon_taf
choice = int(input("Masukkan pilihan data yang akan direkon (1: Rekon ACC, 2: Rekon FIF, 3: Rekon TAF, 4: Rekon Amartha, 5: Rekon EASYCASH, 6: All): "))

match choice:
    case 1:
        rekon_acc()
    case 2:
        rekon_fif()
    case 3:
        rekon_taf()
    case 4:
        rekon_amartha()
    case 5:
        rekon_easycash()
    case 6:
        rekon_acc()
        rekon_fif()
        rekon_taf()
        rekon_amartha()
        rekon_easycash()
    case _:
        print("Pilihan tidak valid.")