import pandas as pd
import numpy as np
import os
from tqdm import tqdm

def rekon_acc():

    total_steps = 15
    with tqdm(total=total_steps, desc="Proses Rekon ACC", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_ACC.xlsx")
        os_acc = pd.read_excel(os.path.join(base_dir, "OS_ACC.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OBJECTCAR_BJJ.xlsx")
        bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca BPKB_ACC.xlsx")
        bpkb_acc = pd.read_excel(os.path.join(base_dir, "BPKB_ACC.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

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
        pbar.set_description("Filter data ACC")
        oss_bjj_acc = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('ASF') | os_bjj['PROD_OFFERING_CODE'].str.startswith('SBSF')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_acc.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_acc['PARTNER_AGRMNT_NO'].astype(str)
        os_acc['PARTNER_AGRMNT_NO'] = os_acc['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan ACC")
        recon_JF_ACC = pd.merge(oss_bjj_acc[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_acc[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_ACC']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_ACC.loc[recon_JF_ACC['_merge'] == 'right_only', 'PARTNER_NAME'] = 'ACC'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_ACC['DIFF_OS'] = recon_JF_ACC['OS_PRINCIPAL_AMT'] - recon_JF_ACC['OS_PRINCIPAL_AMT_ACC']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
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
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
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
        pbar.update(1)

        print(recon_JF_ACC)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_ACC,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_ACC,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_ACC', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_ACC.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_ACC.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")


def rekon_fif():

    total_steps = 16
    with tqdm(total=total_steps, desc="Proses Rekon FIF", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_FIF.xlsx")
        os_fif = pd.read_excel(os.path.join(base_dir, "OS_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OBJECTCAR_BJJ.xlsx")
        bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca BPKB_FIF.xlsx")
        bpkb_fif = pd.read_excel(os.path.join(base_dir, "BPKB_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

        # # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        # os_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/CORACCOUNT_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
        # os_fif = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OS_FIF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
        # bpkb_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OBJECTCAR_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
        # bpkb_fif = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/BPKB_FIF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})

        # # Load data from text files with delimiter '|'
        # inventory_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/ASSETINVENTORY.txt", delimiter="|", low_memory=False)
        # write_off_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/WRITEOFF.txt", delimiter="|", low_memory=False)
        pbar.set_description("Mengubah tipe data OS Principal Amount FIF")
        os_fif['OS_PRINCIPAL_AMT_FIF'] = pd.to_numeric(os_fif['OS_PRINCIPAL_AMT_FIF'], errors='coerce')
        pbar.update(1)

        # Cek Nama Kolom yang ada di tabel
        print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(os_fif.columns) # Periksa nama kolom yang akan digunakan

        # Filter OS BJJ ACC
        pbar.set_description("Filter data FIF")
        oss_bjj_fif = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('FIF') | os_bjj['PROD_OFFERING_CODE'].str.startswith('FIF2')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_fif.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_fif['PARTNER_AGRMNT_NO'].astype(str)
        os_fif['PARTNER_AGRMNT_NO'] = os_fif['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan FIF")
        recon_JF_FIF = pd.merge(oss_bjj_fif[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_fif[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_FIF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_FIF.loc[recon_JF_FIF['_merge'] == 'right_only', 'PARTNER_NAME'] = 'FIF'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_FIF['DIFF_OS'] = recon_JF_FIF['OS_PRINCIPAL_AMT'] - recon_JF_FIF['OS_PRINCIPAL_AMT_FIF']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
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
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
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
        pbar.update(1)

        print(recon_JF_FIF)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_FIF,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_FIF,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_FIF', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_FIF.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_FIF.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")


def rekon_taf():
 
    total_steps = 15
    with tqdm(total=total_steps, desc="Proses Rekon TAF", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_TAF.xlsx")
        os_taf = pd.read_excel(os.path.join(base_dir, "OS_TAF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OBJECTCAR_BJJ.xlsx")
        bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca BPKB_TAF.xlsx")
        bpkb_taf = pd.read_excel(os.path.join(base_dir, "BPKB_TAF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

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
        pbar.set_description("Filter data TAF")
        oss_bjj_taf = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('TAF') | os_bjj['PROD_OFFERING_CODE'].str.startswith('TAF2')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_taf.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_taf['PARTNER_AGRMNT_NO'].astype(str)
        os_taf['PARTNER_AGRMNT_NO'] = os_taf['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan TAF")
        recon_JF_TAF = pd.merge(oss_bjj_taf[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_taf[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_TAF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_TAF.loc[recon_JF_TAF['_merge'] == 'right_only', 'PARTNER_NAME'] = 'TAF'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_TAF['DIFF_OS'] = recon_JF_TAF['OS_PRINCIPAL_AMT'] - recon_JF_TAF['OS_PRINCIPAL_AMT_TAF']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
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
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
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
        pbar.update(1)

        print(recon_JF_TAF)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_TAF,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_TAF,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_TAF', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_TAF.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_TAF.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")

def rekon_amartha():
 
    total_steps = 13
    with tqdm(total=total_steps, desc="Proses Rekon Amartha", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_AMARTHA.xlsx")
        os_amartha = pd.read_excel(os.path.join(base_dir, "OS_AMARTHA.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

        # Cek Nama Kolom yang ada di tabel
        print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(os_amartha.columns) # Periksa nama kolom yang akan digunakan

        # Filter OS BJJ ACC
        pbar.set_description("Filter data Amartha")
        oss_bjj_amartha = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('AMARTHA') | os_bjj['PROD_OFFERING_CODE'].str.startswith('AMARTHA')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_amartha.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_amartha['PARTNER_AGRMNT_NO'].astype(str)
        os_amartha['PARTNER_AGRMNT_NO'] = os_amartha['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan Amartha")
        recon_JF_AMARTHA = pd.merge(oss_bjj_amartha[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_amartha[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_AMARTHA']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_AMARTHA.loc[recon_JF_AMARTHA['_merge'] == 'right_only', 'PARTNER_NAME'] = 'AMARTHA'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_AMARTHA['DIFF_OS'] = recon_JF_AMARTHA['OS_PRINCIPAL_AMT'] - recon_JF_AMARTHA['OS_PRINCIPAL_AMT_AMARTHA']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
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
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
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
        pbar.update(1)

        print(recon_JF_AMARTHA)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_AMARTHA,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_AMARTHA,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_AMARTHA', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_AMARTHA.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_AMARTHA.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")

def rekon_easycash():
 
    total_steps = 13
    with tqdm(total=total_steps, desc="Proses Rekon Easycash", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_EASYCASH.xlsx")
        os_easycash = pd.read_excel(os.path.join(base_dir, "OS_EASYCASH.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

        # Cek Nama Kolom yang ada di tabel
        print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(os_easycash.columns) # Periksa nama kolom yang akan digunakan

        # Filter OS BJJ ACC
        pbar.set_description("Filter data Easycash")
        oss_bjj_easycash = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('EASYCASH') | os_bjj['PROD_OFFERING_CODE'].str.startswith('EASYCASH')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_easycash.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_easycash['PARTNER_AGRMNT_NO'].astype(str)
        os_easycash['PARTNER_AGRMNT_NO'] = os_easycash['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan Easycash")
        recon_JF_EASYCASH = pd.merge(oss_bjj_easycash[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_easycash[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_EASYCASH']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_EASYCASH.loc[recon_JF_EASYCASH['_merge'] == 'right_only', 'PARTNER_NAME'] = 'EASYCASH'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_EASYCASH['DIFF_OS'] = recon_JF_EASYCASH['OS_PRINCIPAL_AMT'] - recon_JF_EASYCASH['OS_PRINCIPAL_AMT_EASYCASH']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
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
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
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
        pbar.update(1)

        print(recon_JF_EASYCASH)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_EASYCASH,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_EASYCASH,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_EASYCASH', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_EASYCASH.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_EASYCASH.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")

def rekon_finture():
 
    total_steps = 13
    with tqdm(total=total_steps, desc="Proses Rekon Finture", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_FINTURE.xlsx")
        os_finture = pd.read_excel(os.path.join(base_dir, "OS_FINTURE.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

        # Cek Nama Kolom yang ada di tabel
        print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(os_finture.columns) # Periksa nama kolom yang akan digunakan

        # Filter OS BJJ ACC
        pbar.set_description("Filter data Finture")
        oss_bjj_finture = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('FINTURE') | os_bjj['PROD_OFFERING_CODE'].str.startswith('FINTURE')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_finture.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_finture['PARTNER_AGRMNT_NO'].astype(str)
        os_finture['PARTNER_AGRMNT_NO'] = os_finture['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan Finture")
        recon_JF_FINTURE = pd.merge(oss_bjj_finture[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_finture[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_FINTURE']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_FINTURE.loc[recon_JF_FINTURE['_merge'] == 'right_only', 'PARTNER_NAME'] = 'FINTURE'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_FINTURE['DIFF_OS'] = recon_JF_FINTURE['OS_PRINCIPAL_AMT'] - recon_JF_FINTURE['OS_PRINCIPAL_AMT_FINTURE']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
        conditions = [
            (recon_JF_FINTURE['_merge'] == 'both'),
            (recon_JF_FINTURE['_merge'] == 'right_only'),
            (recon_JF_FINTURE['_merge'] == 'left_only') & (recon_JF_FINTURE['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
            (recon_JF_FINTURE['_merge'] == 'left_only') & (recon_JF_FINTURE['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
            (recon_JF_FINTURE['_merge'] == 'left_only') & (recon_JF_FINTURE['CONTRACT_STATUS'] == 'EXP')
        ]
        choices = [
            'Data Tersedia di BJJ & Partner',
            'Data hanya Tersedia di FINTURE',
            'Data hanya tersedia di BJJ dan Statusnya Inventory',
            'Data hanya tersedia di BJJ dan Statusnya Writeoff',
            'Data hanya tersedia di BJJ dan Statusnya Lunas'
        ]
        recon_JF_FINTURE['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
        conditions = [
            (recon_JF_FINTURE['DIFF_OS'].isna()),
            (recon_JF_FINTURE['DIFF_OS'] > 100),
            (recon_JF_FINTURE['DIFF_OS'] < -100),
            (recon_JF_FINTURE['DIFF_OS'] > 0) & (recon_JF_FINTURE['DIFF_OS'] <= 100),
            (recon_JF_FINTURE['DIFF_OS'] < 0) & (recon_JF_FINTURE['DIFF_OS'] >= -100),
            (recon_JF_FINTURE['DIFF_OS'] == 0)
        ]
        choices = [
            recon_JF_FINTURE['DATA_STATUS'],
            'Selisih > 100',
            'Selisih < -100',
            'Selisih +- 100',
            'Selisih +- 100',
            'Selisih 0'
        ]
        recon_JF_FINTURE['DIFF_OS_STATUS'] = np.select(conditions, choices)
        pbar.update(1)

        print(recon_JF_FINTURE)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_FINTURE,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_FINTURE,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_FINTURE', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_FINTURE.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_FINTURE.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")

def rekon_hci():
 
    total_steps = 13
    with tqdm(total=total_steps, desc="Proses Rekon HCI", unit=" step", colour="green") as pbar:
        # Mendapatkan direktori tempat script disimpan
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_HCI.xlsx")
        os_hci = pd.read_excel(os.path.join(base_dir, "OS_HCI.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)

        # Load data from text files with delimiter '|'
        pbar.set_description("Membaca ASSETINVENTORY.txt")
        inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
        pbar.update(1)
        pbar.set_description("Membaca WRITEOFF.txt")
        write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)
        pbar.update(1)

        # Cek Nama Kolom yang ada di tabel
        print(os_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(os_hci.columns) # Periksa nama kolom yang akan digunakan

        # Filter OS BJJ ACC
        pbar.set_description("Filter data HCI")
        oss_bjj_hci = os_bjj[os_bjj['PROD_OFFERING_CODE'].str.startswith('HCI') | os_bjj['PROD_OFFERING_CODE'].str.startswith('HCI')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        oss_bjj_hci.loc[:, 'PARTNER_AGRMNT_NO'] = oss_bjj_hci['PARTNER_AGRMNT_NO'].astype(str)
        os_hci['PARTNER_AGRMNT_NO'] = os_hci['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BJJ dan HCI")
        recon_JF_HCI = pd.merge(oss_bjj_hci[['PARTNER_NAME','CUST_NAME','DEFAULT_STATUS','CONTRACT_STATUS','OVERDUE_DAYS','PARTNER_AGRMNT_NO','AGRMNT_NO','OS_PRINCIPAL_AMT']], os_hci[['PARTNER_AGRMNT_NO','OS_PRINCIPAL_AMT_HCI']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        recon_JF_HCI.loc[recon_JF_HCI['_merge'] == 'right_only', 'PARTNER_NAME'] = 'HCI'
        pbar.update(1)

        # Menambahkan kolom DIFF_OS
        pbar.set_description("Menghitung selisih OS")
        recon_JF_HCI['DIFF_OS'] = recon_JF_HCI['OS_PRINCIPAL_AMT'] - recon_JF_HCI['OS_PRINCIPAL_AMT_HCI']
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status data")
        conditions = [
            (recon_JF_HCI['_merge'] == 'both'),
            (recon_JF_HCI['_merge'] == 'right_only'),
            (recon_JF_HCI['_merge'] == 'left_only') & (recon_JF_HCI['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
            (recon_JF_HCI['_merge'] == 'left_only') & (recon_JF_HCI['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
            (recon_JF_HCI['_merge'] == 'left_only') & (recon_JF_HCI['CONTRACT_STATUS'] == 'EXP')
        ]
        choices = [
            'Data Tersedia di BJJ & Partner',
            'Data hanya Tersedia di HCI',
            'Data hanya tersedia di BJJ dan Statusnya Inventory',
            'Data hanya tersedia di BJJ dan Statusnya Writeoff',
            'Data hanya tersedia di BJJ dan Statusnya Lunas'
        ]
        recon_JF_HCI['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')
        pbar.update(1)

        # Menambahkan kolom DIFF_OS_STATUS
        pbar.set_description("Menentukan status selisih OS")
        conditions = [
            (recon_JF_HCI['DIFF_OS'].isna()),
            (recon_JF_HCI['DIFF_OS'] > 100),
            (recon_JF_HCI['DIFF_OS'] < -100),
            (recon_JF_HCI['DIFF_OS'] > 0) & (recon_JF_HCI['DIFF_OS'] <= 100),
            (recon_JF_HCI['DIFF_OS'] < 0) & (recon_JF_HCI['DIFF_OS'] >= -100),
            (recon_JF_HCI['DIFF_OS'] == 0)
        ]
        choices = [
            recon_JF_HCI['DATA_STATUS'],
            'Selisih > 100',
            'Selisih < -100',
            'Selisih +- 100',
            'Selisih +- 100',
            'Selisih 0'
        ]
        recon_JF_HCI['DIFF_OS_STATUS'] = np.select(conditions, choices)
        pbar.update(1)

        print(recon_JF_HCI)

        # # Membuat pivot table Status availibility data
        pbar.set_description("Membuat pivot table status data")
        pivot_table_data = pd.pivot_table(
            recon_JF_HCI,
            values=['PARTNER_AGRMNT_NO'],
            index=['DATA_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_data)

        # # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot table selisih OS")
        pivot_table_diff_os = pd.pivot_table(
            recon_JF_HCI,
            values=['DIFF_OS', 'OS_PRINCIPAL_AMT_HCI', 'PARTNER_AGRMNT_NO'],
            index=['DIFF_OS_STATUS'],
            columns=['PARTNER_NAME'],
            aggfunc={'DIFF_OS': 'sum', 'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)

        print(pivot_table_diff_os)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_os")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_OS_HCI.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_JF_HCI.to_excel(writer, sheet_name='Hasil Recon', index=False)
            pivot_table_data.to_excel(writer, sheet_name='Hasil Pivot data')
            pivot_table_diff_os.to_excel(writer, sheet_name='Hasil Pivot selisih')
        pbar.update(1)
    print(f"File Excel berhasil dibuat di: {output_path}")

# Panggil fungsi rekon_acc, rekon_fif, dan rekon_taf
choice = int(input("Masukkan pilihan data yang akan direkon (1: Rekon ACC, 2: Rekon FIF, 3: Rekon TAF, 4: Rekon Amartha, 5: Rekon EASYCASH, 6: Rekon FINTURE, 7: Rekon HCI, 8: All): "))

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
        rekon_finture()
    case 7:
        rekon_hci()
    case 8:
        rekon_acc()
        rekon_fif()
        rekon_taf()
        rekon_amartha()
        rekon_easycash()
        rekon_finture()
        rekon_hci()

    case _:
        print("Pilihan tidak valid.")