import pandas as pd
import numpy as np
import os
from tqdm import tqdm

def recon_bpkb_acc():

    total_steps = 18
    with tqdm(total=total_steps, desc="Proses Rekon BPKB ACC", unit=" step", colour="green") as pbar:
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

        # Cek Nama Kolom yang ada di tabel
        print(bpkb_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(bpkb_acc.columns) # Periksa nama kolom yang akan digunakan

        # Filter BPKB BJJ ACC
        pbar.set_description("Filter data BPKB ACC")
        bpkb_bjj_acc = bpkb_bjj[bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('ASF') | bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('SBSF')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        bpkb_bjj_acc.loc[:, 'PARTNER_AGRMNT_NO'] = bpkb_bjj_acc['PARTNER_AGRMNT_NO'].astype(str)
        bpkb_acc['PARTNER_AGRMNT_NO'] = bpkb_acc['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BPKB")
        recon_bpkb_JF_ACC = pd.merge(bpkb_bjj_acc[['PARTNER_AGRMNT_NO','AGRMNT_NO','PROD_OFFERING_CODE','ASSET_NAME','BPKB_OWNER','CHASSIS_NO','ENGINE_NO','BPKB_STATUS']], bpkb_acc[['PARTNER_AGRMNT_NO','BPKB_OWNER_ACC','CHASSIS_NO_ACC','ENGINE_NO_ACC','BPKB_STATUS_ACC']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status ketersediaan data")
        conditions = [
            (recon_bpkb_JF_ACC['_merge'] == 'both'),
            (recon_bpkb_JF_ACC['_merge'] == 'right_only'),
            (recon_bpkb_JF_ACC['_merge'] == 'left_only') & (recon_bpkb_JF_ACC['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
            (recon_bpkb_JF_ACC['_merge'] == 'left_only') & (recon_bpkb_JF_ACC['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
            (recon_bpkb_JF_ACC['_merge'] == 'left_only') & (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'RLS')
        ]
        choices = [
            'Data Tersedia di BJJ & Partner',
            'Data hanya Tersedia di ACC',
            'Data hanya tersedia di BJJ dan Statusnya Inventory',
            'Data hanya tersedia di BJJ dan Statusnya Writeoff',
            'Data hanya tersedia di BJJ dan Statusnya Lunas'
        ]
        recon_bpkb_JF_ACC['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')
        pbar.update(1)

        # Menambahkan kolom "REKON STATUS BPKB"
        pbar.set_description("Menentukan status BPKB")
        conditions_bpkb = [
            (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'IN'),
            (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'SP'),
            (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'IN'),
            (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'OUT')
        ]
        choices_bpkb = ['Match', 'Match', 'Need Update Status in Confins','Unmatch']
        recon_bpkb_JF_ACC['REKON_STATUS-BPKB'] = np.select(conditions_bpkb, choices_bpkb, default=recon_bpkb_JF_ACC['DATA_STATUS'])
        pbar.update(1)

        # Menghapus spasi dari kolom CHASSIS_NO dan ENGINE_NO sebelum perbandingan
        pbar.set_description("Membersihkan data nomor rangka & mesin")
        recon_bpkb_JF_ACC['CHASSIS_NO'] = recon_bpkb_JF_ACC['CHASSIS_NO'].str.replace(' ', '')
        recon_bpkb_JF_ACC['CHASSIS_NO_ACC'] = recon_bpkb_JF_ACC['CHASSIS_NO_ACC'].str.replace(' ', '')
        recon_bpkb_JF_ACC['ENGINE_NO'] = recon_bpkb_JF_ACC['ENGINE_NO'].str.replace(' ', '')
        recon_bpkb_JF_ACC['ENGINE_NO_ACC'] = recon_bpkb_JF_ACC['ENGINE_NO_ACC'].str.replace(' ', '')
        pbar.update(1)

        def rangka_status(row):
            a = row['CHASSIS_NO']
            b = row['CHASSIS_NO_ACC']
            if pd.isna(a) or pd.isna(b): return False
            a, b = str(a), str(b)
            if a == b: return True
            elif len(a) == len(b) + 1 and a.endswith('1') and a[:-1] == b: return True
            return False

        pbar.set_description("Menganalisa status data rangka")
        recon_bpkb_JF_ACC['STATUS DATA RANGKA'] = recon_bpkb_JF_ACC.apply(rangka_status, axis=1)
        pbar.update(1)

        def mesin_status(row):
            a = row['ENGINE_NO']
            b = row['ENGINE_NO_ACC']
            if pd.isna(a) or pd.isna(b): return False
            a, b = str(a), str(b)
            if a == b: return True
            elif len(a) == len(b) + 1 and a.startswith('1') and a[1:] == b: return True
            return False

        pbar.set_description("Menganalisa status data mesin")
        recon_bpkb_JF_ACC['STATUS DATA MESIN'] = recon_bpkb_JF_ACC.apply(mesin_status, axis=1)
        pbar.update(1)

        # Menambahkan kolom "STATUS DATA KENDARAAN"
        pbar.set_description("Menentukan status data kendaraan")
        conditions_kendaraan = [
            (recon_bpkb_JF_ACC['BPKB_OWNER_ACC'].isna()) & (recon_bpkb_JF_ACC['CHASSIS_NO_ACC'].isna()) & (recon_bpkb_JF_ACC['ENGINE_NO_ACC'].isna()),
            (recon_bpkb_JF_ACC['STATUS DATA RANGKA'] == True) & (recon_bpkb_JF_ACC['STATUS DATA MESIN'] == True),
            (recon_bpkb_JF_ACC['STATUS DATA RANGKA'] == True) & (recon_bpkb_JF_ACC['STATUS DATA MESIN'] == False),
            (recon_bpkb_JF_ACC['STATUS DATA RANGKA'] == False) & (recon_bpkb_JF_ACC['STATUS DATA MESIN'] == True),
            (recon_bpkb_JF_ACC['STATUS DATA RANGKA'] == False) & (recon_bpkb_JF_ACC['STATUS DATA MESIN'] == False),
        ]
        choices_kendaraan = [
            recon_bpkb_JF_ACC['DATA_STATUS'],
            '1. All Match',
            '2. ENGINE_NO_Unmatch',
            '3. CHASSIS_NO_Unmatch',
            '4. All Unmatch'
        ]
        recon_bpkb_JF_ACC['STATUS DATA KENDARAAN'] = np.select(conditions_kendaraan, choices_kendaraan, default=recon_bpkb_JF_ACC['DATA_STATUS'])
        pbar.update(1)

        print(recon_bpkb_JF_ACC)

        # Membuat pivot table REKON_STATUS-BPKB
        pbar.set_description("Membuat pivot status BPKB")
        pivot_table_data_status_bpkb = pd.pivot_table(
            recon_bpkb_JF_ACC,
            values=[ 'PARTNER_AGRMNT_NO'],
            index=['REKON_STATUS-BPKB'],
            columns=['PROD_OFFERING_CODE'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)
        print(pivot_table_data_status_bpkb)

        # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot status kendaraan")
        pivot_table_status_data_kendaraan = pd.pivot_table(
            recon_bpkb_JF_ACC,
            values=[ 'PARTNER_AGRMNT_NO'],
            index=['STATUS DATA KENDARAAN'],
            columns=['PROD_OFFERING_CODE'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)
        print(pivot_table_status_data_kendaraan)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_bpkb")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_bpkb_ACC.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_bpkb_JF_ACC.to_excel(writer, sheet_name='Hasil Recon BPKB', index=False)
            pivot_table_data_status_bpkb.to_excel(writer, sheet_name='Hasil Pivot data Status BPKB')
            pivot_table_status_data_kendaraan.to_excel(writer, sheet_name='Hasil Pivot data Kendaraan')
        pbar.update(1)
        print(f"File Excel berhasil dibuat di: {output_path}")

def recon_bpkb_fif():



    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))
    total_steps = 16
    with tqdm(total=total_steps, desc="Proses Rekon BPKB FIF", unit=" step", colour="green") as pbar:
        # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
        pbar.set_description("Membaca CORACCOUNT_BJJ.xlsx")
        os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
        pbar.update(1)
        pbar.set_description("Membaca OS_FIF.xlsx")
        os_acc = pd.read_excel(os.path.join(base_dir, "OS_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
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

        # Cek Nama Kolom yang ada di tabel
        print(bpkb_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(bpkb_fif.columns) # Periksa nama kolom yang akan digunakan

        # Filter BPKB BJJ ACC
        pbar.set_description("Filter data BPKB FIF")
        bpkb_bjj_fif = bpkb_bjj[bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('FIF') | bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('FIF2')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        bpkb_bjj_fif.loc[:, 'PARTNER_AGRMNT_NO'] = bpkb_bjj_fif['PARTNER_AGRMNT_NO'].astype(str)
        bpkb_fif['PARTNER_AGRMNT_NO'] = bpkb_fif['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BPKB")
        recon_bpkb_JF_FIF = pd.merge(bpkb_bjj_fif[['PARTNER_AGRMNT_NO','AGRMNT_NO','PROD_OFFERING_CODE','ASSET_NAME','BPKB_OWNER','CHASSIS_NO','ENGINE_NO','BPKB_STATUS']], bpkb_fif[['PARTNER_AGRMNT_NO','BPKB_OWNER_FIF','CHASSIS_NO_FIF','ENGINE_NO_FIF','BPKB_STATUS_FIF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status ketersediaan data")
        conditions = [
            (recon_bpkb_JF_FIF['_merge'] == 'both'),
            (recon_bpkb_JF_FIF['_merge'] == 'right_only'),
            (recon_bpkb_JF_FIF['_merge'] == 'left_only') & (recon_bpkb_JF_FIF['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
            (recon_bpkb_JF_FIF['_merge'] == 'left_only') & (recon_bpkb_JF_FIF['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
            (recon_bpkb_JF_FIF['_merge'] == 'left_only') & (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'RLS')
        ]
        choices = [
            'Data Tersedia di BJJ & Partner',
            'Data hanya Tersedia di FIF',
            'Data hanya tersedia di BJJ dan Statusnya Inventory',
            'Data hanya tersedia di BJJ dan Statusnya Writeoff',
            'Data hanya tersedia di BJJ dan Statusnya Lunas'
        ]
        recon_bpkb_JF_FIF['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')
        pbar.update(1)

        # Menambahkan kolom "REKON STATUS BPKB"
        pbar.set_description("Menentukan status BPKB")
        conditions_bpkb = [
            (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'OHD'),
            (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'TBO'),
            (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'OHD'),
            (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'TBO'),
            (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'RLS') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'OHD')
        ]
        choices_bpkb = ['Match', 'Match', 'Need Update Status in Confins','Unmatch','Match']
        recon_bpkb_JF_FIF['REKON_STATUS-BPKB'] = np.select(conditions_bpkb, choices_bpkb, default=recon_bpkb_JF_FIF['DATA_STATUS'])
        pbar.update(1)

        # Menghapus spasi dari kolom CHASSIS_NO dan ENGINE_NO sebelum perbandingan
        pbar.set_description("Membersihkan data nomor rangka & mesin")
        recon_bpkb_JF_FIF['CHASSIS_NO'] = recon_bpkb_JF_FIF['CHASSIS_NO'].str.replace(' ', '')
        recon_bpkb_JF_FIF['CHASSIS_NO_FIF'] = recon_bpkb_JF_FIF['CHASSIS_NO_FIF'].str.replace(' ', '')
        recon_bpkb_JF_FIF['ENGINE_NO'] = recon_bpkb_JF_FIF['ENGINE_NO'].str.replace(' ', '')
        recon_bpkb_JF_FIF['ENGINE_NO_FIF'] = recon_bpkb_JF_FIF['ENGINE_NO_FIF'].str.replace(' ', '')
        pbar.update(1)

        # Menambahkan kolom "STATUS DATA KENDARAAN"
        pbar.set_description("Menentukan status data kendaraan")
        conditions_kendaraan = [
            (recon_bpkb_JF_FIF['BPKB_OWNER_FIF'].isna()) & (recon_bpkb_JF_FIF['CHASSIS_NO_FIF'].isna()) & (recon_bpkb_JF_FIF['ENGINE_NO_FIF'].isna()),
            (recon_bpkb_JF_FIF['CHASSIS_NO'] == recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] == recon_bpkb_JF_FIF['ENGINE_NO_FIF']),
            (recon_bpkb_JF_FIF['CHASSIS_NO'] == recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] != recon_bpkb_JF_FIF['ENGINE_NO_FIF']),
            (recon_bpkb_JF_FIF['CHASSIS_NO'] != recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] == recon_bpkb_JF_FIF['ENGINE_NO_FIF']),
            (recon_bpkb_JF_FIF['CHASSIS_NO'] != recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] != recon_bpkb_JF_FIF['ENGINE_NO_FIF'])
        ]
        choices_kendaraan = [recon_bpkb_JF_FIF['DATA_STATUS'],'All Match', 'ENGINE_NO_Unmatch', 'CHASSIS_NO_Unmatch','All Unmatch']
        recon_bpkb_JF_FIF['STATUS DATA KENDARAAN'] = np.select(conditions_kendaraan, choices_kendaraan, default=recon_bpkb_JF_FIF['DATA_STATUS'])
        pbar.update(1)

        print(recon_bpkb_JF_FIF)

        # Membuat pivot table REKON_STATUS-BPKB
        pbar.set_description("Membuat pivot status BPKB")
        pivot_table_data_status_bpkb = pd.pivot_table(
            recon_bpkb_JF_FIF,
            values=[ 'PARTNER_AGRMNT_NO'],
            index=['REKON_STATUS-BPKB'],
            columns=['PROD_OFFERING_CODE'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)
        print(pivot_table_data_status_bpkb)

        # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot status kendaraan")
        pivot_table_status_data_kendaraan = pd.pivot_table(
            recon_bpkb_JF_FIF,
            values=[ 'PARTNER_AGRMNT_NO'],
            index=['STATUS DATA KENDARAAN'],
            columns=['PROD_OFFERING_CODE'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)
        print(pivot_table_status_data_kendaraan)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_bpkb")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_bpkb_FIF.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_bpkb_JF_FIF.to_excel(writer, sheet_name='Hasil Recon BPKB', index=False)
            pivot_table_data_status_bpkb.to_excel(writer, sheet_name='Hasil Pivot data Status BPKB')
            pivot_table_status_data_kendaraan.to_excel(writer, sheet_name='Hasil Pivot data Kendaraan')
        pbar.update(1)
        print(f"File Excel berhasil dibuat di: {output_path}")

def recon_bpkb_taf():
    
    
    total_steps = 16
    with tqdm(total=total_steps, desc="Proses Rekon BPKB TAF", unit=" step", colour="green") as pbar:
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

        # Cek Nama Kolom yang ada di tabel
        print(bpkb_bjj.columns) # Periksa nama kolom yang akan digunakan
        print(bpkb_taf.columns) # Periksa nama kolom yang akan digunakan

        # Filter BPKB BJJ ACC
        pbar.set_description("Filter data BPKB TAF")
        bpkb_bjj_taf = bpkb_bjj[bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('TAF') | bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('TAF2')]
        pbar.update(1)

        # Konversi type data agar sama dengan coreaccount
        pbar.set_description("Konversi tipe data")
        bpkb_bjj_taf.loc[:, 'PARTNER_AGRMNT_NO'] = bpkb_bjj_taf['PARTNER_AGRMNT_NO'].astype(str)
        bpkb_taf['PARTNER_AGRMNT_NO'] = bpkb_taf['PARTNER_AGRMNT_NO'].astype(str)
        pbar.update(1)

        pbar.set_description("Menggabungkan data BPKB")
        recon_bpkb_JF_TAF = pd.merge(bpkb_bjj_taf[['PARTNER_AGRMNT_NO','AGRMNT_NO','PROD_OFFERING_CODE','ASSET_NAME','BPKB_OWNER','CHASSIS_NO','ENGINE_NO','BPKB_STATUS']], bpkb_taf[['PARTNER_AGRMNT_NO','BPKB_OWNER_TAF','CHASSIS_NO_TAF','ENGINE_NO_TAF','BPKB_STATUS_TAF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)
        pbar.update(1)

        # Menambahkan kolom DATA_STATUS
        pbar.set_description("Menentukan status ketersediaan data")
        conditions = [
            (recon_bpkb_JF_TAF['_merge'] == 'both'),
            (recon_bpkb_JF_TAF['_merge'] == 'right_only'),
            (recon_bpkb_JF_TAF['_merge'] == 'left_only') & (recon_bpkb_JF_TAF['AGRMNT_NO'].isin(inventory_df['AGRMNT_NO'])),
            (recon_bpkb_JF_TAF['_merge'] == 'left_only') & (recon_bpkb_JF_TAF['AGRMNT_NO'].isin(write_off_df['AGRMNT_NO'])),
            (recon_bpkb_JF_TAF['_merge'] == 'left_only') & (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'RLS')
        ]
        choices = [
            'Data Tersedia di BJJ & Partner',
            'Data hanya Tersedia di TAF',
            'Data hanya tersedia di BJJ dan Statusnya Inventory',
            'Data hanya tersedia di BJJ dan Statusnya Writeoff',
            'Data hanya tersedia di BJJ dan Statusnya Lunas'
        ]
        recon_bpkb_JF_TAF['DATA_STATUS'] = np.select(conditions, choices, default='Need Confirmation')
        pbar.update(1)

        # Menambahkan kolom "REKON STATUS BPKB"
        pbar.set_description("Menentukan status BPKB")
        conditions_bpkb = [
            (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'RLS') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'IN'),
            (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'IN'),
            (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'SP'),
            (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'IN'),
            (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'OUT')
        ]
        choices_bpkb = ['Match','Match', 'Match', 'Need Update Status in Confins','Unmatch']
        recon_bpkb_JF_TAF['REKON_STATUS-BPKB'] = np.select(conditions_bpkb, choices_bpkb, default=recon_bpkb_JF_TAF['DATA_STATUS'])
        pbar.update(1)

        # Menghapus spasi dari kolom CHASSIS_NO dan ENGINE_NO sebelum perbandingan
        pbar.set_description("Membersihkan data nomor rangka & mesin")
        recon_bpkb_JF_TAF['CHASSIS_NO'] = recon_bpkb_JF_TAF['CHASSIS_NO'].str.replace(' ', '')
        recon_bpkb_JF_TAF['CHASSIS_NO_TAF'] = recon_bpkb_JF_TAF['CHASSIS_NO_TAF'].str.replace(' ', '')
        recon_bpkb_JF_TAF['ENGINE_NO'] = recon_bpkb_JF_TAF['ENGINE_NO'].str.replace(' ', '')
        recon_bpkb_JF_TAF['ENGINE_NO_TAF'] = recon_bpkb_JF_TAF['ENGINE_NO_TAF'].str.replace(' ', '')
        pbar.update(1)

        # Menambahkan kolom "STATUS DATA KENDARAAN"
        pbar.set_description("Menentukan status data kendaraan")
        conditions_kendaraan = [
            (recon_bpkb_JF_TAF['BPKB_OWNER_TAF'].isna()) & (recon_bpkb_JF_TAF['CHASSIS_NO_TAF'].isna()) & (recon_bpkb_JF_TAF['ENGINE_NO_TAF'].isna()),
            (recon_bpkb_JF_TAF['CHASSIS_NO'] == recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] == recon_bpkb_JF_TAF['ENGINE_NO_TAF']),
            (recon_bpkb_JF_TAF['CHASSIS_NO'] == recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] != recon_bpkb_JF_TAF['ENGINE_NO_TAF']),
            (recon_bpkb_JF_TAF['CHASSIS_NO'] != recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] == recon_bpkb_JF_TAF['ENGINE_NO_TAF']),
            (recon_bpkb_JF_TAF['CHASSIS_NO'] != recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] != recon_bpkb_JF_TAF['ENGINE_NO_TAF'])
        ]
        choices_kendaraan = [recon_bpkb_JF_TAF['DATA_STATUS'],'All Match', 'ENGINE_NO_Unmatch', 'CHASSIS_NO_Unmatch','All Unmatch']
        recon_bpkb_JF_TAF['STATUS DATA KENDARAAN'] = np.select(conditions_kendaraan, choices_kendaraan, default=recon_bpkb_JF_TAF['DATA_STATUS'])
        pbar.update(1)

        print(recon_bpkb_JF_TAF)

        # Membuat pivot table REKON_STATUS-BPKB
        pbar.set_description("Membuat pivot status BPKB")
        pivot_table_data_status_bpkb = pd.pivot_table(
            recon_bpkb_JF_TAF,
            values=[ 'PARTNER_AGRMNT_NO'],
            index=['REKON_STATUS-BPKB'],
            columns=['PROD_OFFERING_CODE'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)
        print(pivot_table_data_status_bpkb)

        # Membuat pivot table selisih OS dan OS Principal
        pbar.set_description("Membuat pivot status kendaraan")
        pivot_table_status_data_kendaraan = pd.pivot_table(
            recon_bpkb_JF_TAF,
            values=[ 'PARTNER_AGRMNT_NO'],
            index=['STATUS DATA KENDARAAN'],
            columns=['PROD_OFFERING_CODE'],
            aggfunc={'PARTNER_AGRMNT_NO': 'count'},
            fill_value=0
        ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})
        pbar.update(1)
        print(pivot_table_status_data_kendaraan)

        # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
        pbar.set_description("Menyimpan hasil ke Excel")
        output_dir = os.path.join(base_dir, "output_bpkb")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "reconciliation_bpkb_TAF.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            recon_bpkb_JF_TAF.to_excel(writer, sheet_name='Hasil Recon BPKB', index=False)
            pivot_table_data_status_bpkb.to_excel(writer, sheet_name='Hasil Pivot data Status BPKB')
            pivot_table_status_data_kendaraan.to_excel(writer, sheet_name='Hasil Pivot data Kendaraan')
        pbar.update(1)
        print(f"File Excel berhasil dibuat di: {output_path}")



# Panggil fungsi rekon_acc, rekon_fif, dan rekon_taf
choice = int(input("Masukkan pilihan data BPKB yang akan direkon (1: Rekon ACC, 2: Rekon FIF, 3: Rekon TAF, 4: Semua): "))

match choice:
    case 1:
        recon_bpkb_acc()
    case 2:
        recon_bpkb_fif()
    case 3:
        recon_bpkb_taf()
    case 4:
        recon_bpkb_acc()
        recon_bpkb_fif()
        recon_bpkb_taf()
    case _:
        print("Pilihan tidak valid.")