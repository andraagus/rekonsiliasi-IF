import pandas as pd
import numpy as np
import os

def recon_bpkb_acc():

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
    print(bpkb_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(bpkb_acc.columns) # Periksa nama kolom yang akan digunakan

    # Filter BPKB BJJ ACC
    bpkb_bjj_acc = bpkb_bjj[bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('ASF') | bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('SBSF')]

    # Konversi type data agar sama dengan coreaccount
    bpkb_bjj_acc.loc[:, 'PARTNER_AGRMNT_NO'] = bpkb_bjj_acc['PARTNER_AGRMNT_NO'].astype(str)
    bpkb_acc['PARTNER_AGRMNT_NO'] = bpkb_acc['PARTNER_AGRMNT_NO'].astype(str)

    recon_bpkb_JF_ACC = pd.merge(bpkb_bjj_acc[['PARTNER_AGRMNT_NO','AGRMNT_NO','PROD_OFFERING_CODE','ASSET_NAME','BPKB_OWNER','CHASSIS_NO','ENGINE_NO','BPKB_STATUS']], bpkb_acc[['PARTNER_AGRMNT_NO','BPKB_OWNER_ACC','CHASSIS_NO_ACC','ENGINE_NO_ACC','BPKB_STATUS_ACC']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DATA_STATUS
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

    # Menambahkan kolom "REKON STATUS BPKB"
    conditions_bpkb = [
        (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'IN'),
        (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'SP'),
        (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'IN'),
        (recon_bpkb_JF_ACC['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_ACC['BPKB_STATUS_ACC'] == 'OUT')
    ]
    choices_bpkb = ['Match', 'Match', 'Need Update Status in Confins','Unmatch']
    recon_bpkb_JF_ACC['REKON_STATUS-BPKB'] = np.select(conditions_bpkb, choices_bpkb, default=recon_bpkb_JF_ACC['DATA_STATUS'])

    # Menghapus spasi dari kolom CHASSIS_NO dan ENGINE_NO sebelum perbandingan
    recon_bpkb_JF_ACC['CHASSIS_NO'] = recon_bpkb_JF_ACC['CHASSIS_NO'].str.replace(' ', '')
    recon_bpkb_JF_ACC['CHASSIS_NO_ACC'] = recon_bpkb_JF_ACC['CHASSIS_NO_ACC'].str.replace(' ', '')
    recon_bpkb_JF_ACC['ENGINE_NO'] = recon_bpkb_JF_ACC['ENGINE_NO'].str.replace(' ', '')
    recon_bpkb_JF_ACC['ENGINE_NO_ACC'] = recon_bpkb_JF_ACC['ENGINE_NO_ACC'].str.replace(' ', '')

    def rangka_status(row):
        """
        Memeriksa apakah string a dan b sesuai dengan kondisi:
        1. Jika a sama dengan b, hasilnya True.
        2. Jika a tidak sama dengan b tetapi bedanya hanya satu karakter di belakang dan itu '1', hasilnya True.
        3. Selain itu, hasilnya False.
        """
        a = row['CHASSIS_NO']
        b = row['CHASSIS_NO_ACC']
        
        # Handle NaN values
        if pd.isna(a) or pd.isna(b):
            return False
            
        # Convert to string if not already
        a = str(a)
        b = str(b)
        
        if a == b:
            return True
        elif len(a) == len(b) + 1 and a.endswith('1') and a[:-1] == b:
            return True
        return False

    recon_bpkb_JF_ACC['STATUS DATA RANGKA'] = recon_bpkb_JF_ACC.apply(rangka_status, axis=1)

    def mesin_status(row):
        """
        Memeriksa apakah string a dan b sesuai dengan kondisi:
        1. Jika a sama dengan b, hasilnya True.
        2. Jika a tidak sama dengan b tetapi bedanya hanya satu karakter di belakang dan itu '1', hasilnya True.
        3. Selain itu, hasilnya False.
        """
        a = row['ENGINE_NO']
        b = row['ENGINE_NO_ACC']
        
        # Handle NaN values
        if pd.isna(a) or pd.isna(b):
            return False
            
        # Convert to string if not already
        a = str(a)
        b = str(b)
        
        if a == b:
            return True
        elif len(a) == len(b) + 1 and a.startswith('1') and a[1:] == b:
            return True
        return False

    recon_bpkb_JF_ACC['STATUS DATA MESIN'] = recon_bpkb_JF_ACC.apply(mesin_status, axis=1)

    # Menambahkan kolom "STATUS DATA KENDARAAN"
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



    print(recon_bpkb_JF_ACC)

    # Membuat pivot table REKON_STATUS-BPKB
    pivot_table_data_status_bpkb = pd.pivot_table(
        recon_bpkb_JF_ACC,
        values=[ 'PARTNER_AGRMNT_NO'],
        index=['REKON_STATUS-BPKB'],
        columns=['PROD_OFFERING_CODE'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data_status_bpkb)


    # Membuat pivot table selisih OS dan OS Principal
    pivot_table_status_data_kendaraan = pd.pivot_table(
        recon_bpkb_JF_ACC,
        values=[ 'PARTNER_AGRMNT_NO'],
        index=['STATUS DATA KENDARAAN'],
        columns=['PROD_OFFERING_CODE'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_status_data_kendaraan)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_bpkb_ACC.xlsx") as writer:
        recon_bpkb_JF_ACC.to_excel(writer, sheet_name='Hasil Recon BPKB', index=False)
        pivot_table_data_status_bpkb.to_excel(writer, sheet_name='Hasil Pivot data Status BPKB')
        pivot_table_status_data_kendaraan.to_excel(writer, sheet_name='Hasil Pivot data Kendaraan')
    print("File Excel berhasil dibuat dengan tiga sheet.")

def recon_bpkb_fif():



    # Mendapatkan direktori tempat script disimpan
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    os_bjj = pd.read_excel(os.path.join(base_dir, "CORACCOUNT_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    os_acc = pd.read_excel(os.path.join(base_dir, "OS_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_bjj = pd.read_excel(os.path.join(base_dir, "OBJECTCAR_BJJ.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})
    bpkb_fif = pd.read_excel(os.path.join(base_dir, "BPKB_FIF.xlsx"), dtype={'PARTNER_AGRMNT_NO': str})

    # Load data from text files with delimiter '|'
    inventory_df = pd.read_csv(os.path.join(base_dir, "ASSETINVENTORY.txt"), delimiter="|", low_memory=False)
    write_off_df = pd.read_csv(os.path.join(base_dir, "WRITEOFF.txt"), delimiter="|", low_memory=False)


    #     # Membaca file Excel dengan memastikan kolom PARTNER_AGRMNT_NO dibaca sebagai string
    # os_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/CORACCOUNT_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # os_acc = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OS_FIF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_bjj = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/OBJECTCAR_BJJ.xlsx", dtype={'PARTNER_AGRMNT_NO': str})
    # bpkb_fif = pd.read_excel("C:/Users/agus.andra/Documents/REKONSILIASI JF/BPKB_FIF.xlsx", dtype={'PARTNER_AGRMNT_NO': str})

    # # Load data from text files with delimiter '|'
    # inventory_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/ASSETINVENTORY.txt", delimiter="|", low_memory=False)
    # write_off_df = pd.read_csv("C:/Users/agus.andra/Documents/REKONSILIASI JF/WRITEOFF.txt", delimiter="|", low_memory=False)

    # Cek Nama Kolom yang ada di tabel
    print(bpkb_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(bpkb_fif.columns) # Periksa nama kolom yang akan digunakan

    # Filter BPKB BJJ ACC
    bpkb_bjj_fif = bpkb_bjj[bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('FIF') | bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('FIF2')]

    # Konversi type data agar sama dengan coreaccount
    bpkb_bjj_fif.loc[:, 'PARTNER_AGRMNT_NO'] = bpkb_bjj_fif['PARTNER_AGRMNT_NO'].astype(str)
    bpkb_fif['PARTNER_AGRMNT_NO'] = bpkb_fif['PARTNER_AGRMNT_NO'].astype(str)

    recon_bpkb_JF_FIF = pd.merge(bpkb_bjj_fif[['PARTNER_AGRMNT_NO','AGRMNT_NO','PROD_OFFERING_CODE','ASSET_NAME','BPKB_OWNER','CHASSIS_NO','ENGINE_NO','BPKB_STATUS']], bpkb_fif[['PARTNER_AGRMNT_NO','BPKB_OWNER_FIF','CHASSIS_NO_FIF','ENGINE_NO_FIF','BPKB_STATUS_FIF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DATA_STATUS
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

    # Menambahkan kolom "REKON STATUS BPKB"
    conditions_bpkb = [
        (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'OHD'),
        (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'TBO'),
        (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'OHD'),
        (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'TBO'),
        (recon_bpkb_JF_FIF['BPKB_STATUS'] == 'RLS') & (recon_bpkb_JF_FIF['BPKB_STATUS_FIF'] == 'OHD')
    ]
    choices_bpkb = ['Match', 'Match', 'Need Update Status in Confins','Unmatch','Match']
    recon_bpkb_JF_FIF['REKON_STATUS-BPKB'] = np.select(conditions_bpkb, choices_bpkb, default=recon_bpkb_JF_FIF['DATA_STATUS'])

    # Menghapus spasi dari kolom CHASSIS_NO dan ENGINE_NO sebelum perbandingan
    recon_bpkb_JF_FIF['CHASSIS_NO'] = recon_bpkb_JF_FIF['CHASSIS_NO'].str.replace(' ', '')
    recon_bpkb_JF_FIF['CHASSIS_NO_FIF'] = recon_bpkb_JF_FIF['CHASSIS_NO_FIF'].str.replace(' ', '')
    recon_bpkb_JF_FIF['ENGINE_NO'] = recon_bpkb_JF_FIF['ENGINE_NO'].str.replace(' ', '')
    recon_bpkb_JF_FIF['ENGINE_NO_FIF'] = recon_bpkb_JF_FIF['ENGINE_NO_FIF'].str.replace(' ', '')

    # Menambahkan kolom "STATUS DATA KENDARAAN"
    conditions_kendaraan = [
        (recon_bpkb_JF_FIF['BPKB_OWNER_FIF'].isna()) & (recon_bpkb_JF_FIF['CHASSIS_NO_FIF'].isna()) & (recon_bpkb_JF_FIF['ENGINE_NO_FIF'].isna()),
        (recon_bpkb_JF_FIF['CHASSIS_NO'] == recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] == recon_bpkb_JF_FIF['ENGINE_NO_FIF']),
        (recon_bpkb_JF_FIF['CHASSIS_NO'] == recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] != recon_bpkb_JF_FIF['ENGINE_NO_FIF']),
        (recon_bpkb_JF_FIF['CHASSIS_NO'] != recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] == recon_bpkb_JF_FIF['ENGINE_NO_FIF']),
        (recon_bpkb_JF_FIF['CHASSIS_NO'] != recon_bpkb_JF_FIF['CHASSIS_NO_FIF']) & (recon_bpkb_JF_FIF['ENGINE_NO'] != recon_bpkb_JF_FIF['ENGINE_NO_FIF'])
    ]
    choices_kendaraan = [recon_bpkb_JF_FIF['DATA_STATUS'],'All Match', 'ENGINE_NO_Unmatch', 'CHASSIS_NO_Unmatch','All Unmatch']
    recon_bpkb_JF_FIF['STATUS DATA KENDARAAN'] = np.select(conditions_kendaraan, choices_kendaraan, default=recon_bpkb_JF_FIF['DATA_STATUS'])

    print(recon_bpkb_JF_FIF)

    # Membuat pivot table REKON_STATUS-BPKB
    pivot_table_data_status_bpkb = pd.pivot_table(
        recon_bpkb_JF_FIF,
        values=[ 'PARTNER_AGRMNT_NO'],
        index=['REKON_STATUS-BPKB'],
        columns=['PROD_OFFERING_CODE'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data_status_bpkb)


    # Membuat pivot table selisih OS dan OS Principal
    pivot_table_status_data_kendaraan = pd.pivot_table(
        recon_bpkb_JF_FIF,
        values=[ 'PARTNER_AGRMNT_NO'],
        index=['STATUS DATA KENDARAAN'],
        columns=['PROD_OFFERING_CODE'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_status_data_kendaraan)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_bpkb__FIF.xlsx") as writer:
        recon_bpkb_JF_FIF.to_excel(writer, sheet_name='Hasil Recon BPKB', index=False)
        pivot_table_data_status_bpkb.to_excel(writer, sheet_name='Hasil Pivot data Status BPKB')
        pivot_table_status_data_kendaraan.to_excel(writer, sheet_name='Hasil Pivot data Kendaraan')
    print("File Excel berhasil dibuat dengan tiga sheet.")

def recon_bpkb_taf():
    
    
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
    print(bpkb_bjj.columns) # Periksa nama kolom yang akan digunakan
    print(bpkb_taf.columns) # Periksa nama kolom yang akan digunakan

    # Filter BPKB BJJ ACC
    bpkb_bjj_taf = bpkb_bjj[bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('TAF') | bpkb_bjj['PROD_OFFERING_CODE'].str.startswith('TAF2')]

    # Konversi type data agar sama dengan coreaccount
    bpkb_bjj_taf.loc[:, 'PARTNER_AGRMNT_NO'] = bpkb_bjj_taf['PARTNER_AGRMNT_NO'].astype(str)
    bpkb_taf['PARTNER_AGRMNT_NO'] = bpkb_taf['PARTNER_AGRMNT_NO'].astype(str)

    recon_bpkb_JF_TAF = pd.merge(bpkb_bjj_taf[['PARTNER_AGRMNT_NO','AGRMNT_NO','PROD_OFFERING_CODE','ASSET_NAME','BPKB_OWNER','CHASSIS_NO','ENGINE_NO','BPKB_STATUS']], bpkb_taf[['PARTNER_AGRMNT_NO','BPKB_OWNER_TAF','CHASSIS_NO_TAF','ENGINE_NO_TAF','BPKB_STATUS_TAF']], on='PARTNER_AGRMNT_NO', how='outer', indicator=True)

    # Menambahkan kolom DATA_STATUS
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

    # Menambahkan kolom "REKON STATUS BPKB"
    conditions_bpkb = [
        (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'RLS') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'IN'),
        (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'IN'),
        (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'SP'),
        (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'WTG') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'IN'),
        (recon_bpkb_JF_TAF['BPKB_STATUS'] == 'ONH') & (recon_bpkb_JF_TAF['BPKB_STATUS_TAF'] == 'OUT')
    ]
    choices_bpkb = ['Match','Match', 'Match', 'Need Update Status in Confins','Unmatch']
    recon_bpkb_JF_TAF['REKON_STATUS-BPKB'] = np.select(conditions_bpkb, choices_bpkb, default=recon_bpkb_JF_TAF['DATA_STATUS'])

    # Menghapus spasi dari kolom CHASSIS_NO dan ENGINE_NO sebelum perbandingan
    recon_bpkb_JF_TAF['CHASSIS_NO'] = recon_bpkb_JF_TAF['CHASSIS_NO'].str.replace(' ', '')
    recon_bpkb_JF_TAF['CHASSIS_NO_TAF'] = recon_bpkb_JF_TAF['CHASSIS_NO_TAF'].str.replace(' ', '')
    recon_bpkb_JF_TAF['ENGINE_NO'] = recon_bpkb_JF_TAF['ENGINE_NO'].str.replace(' ', '')
    recon_bpkb_JF_TAF['ENGINE_NO_TAF'] = recon_bpkb_JF_TAF['ENGINE_NO_TAF'].str.replace(' ', '')

    # Menambahkan kolom "STATUS DATA KENDARAAN"
    conditions_kendaraan = [
        (recon_bpkb_JF_TAF['BPKB_OWNER_TAF'].isna()) & (recon_bpkb_JF_TAF['CHASSIS_NO_TAF'].isna()) & (recon_bpkb_JF_TAF['ENGINE_NO_TAF'].isna()),
        (recon_bpkb_JF_TAF['CHASSIS_NO'] == recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] == recon_bpkb_JF_TAF['ENGINE_NO_TAF']),
        (recon_bpkb_JF_TAF['CHASSIS_NO'] == recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] != recon_bpkb_JF_TAF['ENGINE_NO_TAF']),
        (recon_bpkb_JF_TAF['CHASSIS_NO'] != recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] == recon_bpkb_JF_TAF['ENGINE_NO_TAF']),
        (recon_bpkb_JF_TAF['CHASSIS_NO'] != recon_bpkb_JF_TAF['CHASSIS_NO_TAF']) & (recon_bpkb_JF_TAF['ENGINE_NO'] != recon_bpkb_JF_TAF['ENGINE_NO_TAF'])
    ]
    choices_kendaraan = [recon_bpkb_JF_TAF['DATA_STATUS'],'All Match', 'ENGINE_NO_Unmatch', 'CHASSIS_NO_Unmatch','All Unmatch']
    recon_bpkb_JF_TAF['STATUS DATA KENDARAAN'] = np.select(conditions_kendaraan, choices_kendaraan, default=recon_bpkb_JF_TAF['DATA_STATUS'])

    print(recon_bpkb_JF_TAF)

    # Kondisi 1: Semua kolom BPKB_OWNER_TAF, CHASSIS_NO_TAF, dan ENGINE_NO_TAF kosong (NaN)
    # Kondisi 2: CHASSIS_NO dan ENGINE_NO cocok dengan CHASSIS_NO_TAF dan ENGINE_NO_TAF
    # Kondisi 3: CHASSIS_NO cocok, tetapi ENGINE_NO tidak cocok
    # Kondisi 4: ENGINE_NO cocok, tetapi CHASSIS_NO tidak cocok
    # Kondisi 5: Baik CHASSIS_NO maupun ENGINE_NO tidak cocok

    # Membuat pivot table REKON_STATUS-BPKB
    pivot_table_data_status_bpkb = pd.pivot_table(
        recon_bpkb_JF_TAF,
        values=[ 'PARTNER_AGRMNT_NO'],
        index=['REKON_STATUS-BPKB'],
        columns=['PROD_OFFERING_CODE'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_data_status_bpkb)


    # Membuat pivot table selisih OS dan OS Principal
    pivot_table_status_data_kendaraan = pd.pivot_table(
        recon_bpkb_JF_TAF,
        values=[ 'PARTNER_AGRMNT_NO'],
        index=['STATUS DATA KENDARAAN'],
        columns=['PROD_OFFERING_CODE'],
        aggfunc={'PARTNER_AGRMNT_NO': 'count'},
        fill_value=0
    ).rename(columns={'PARTNER_AGRMNT_NO': 'Jumlah Data'})

    print(pivot_table_status_data_kendaraan)

    # Menulis hasil recon dan pivot ke dalam satu file Excel dengan sheet yang berbeda
    with pd.ExcelWriter("C:/Users/agus.andra/Documents/REKONSILIASI JF/reconciliation_bpkb_TAF.xlsx") as writer:
        recon_bpkb_JF_TAF.to_excel(writer, sheet_name='Hasil Recon BPKB', index=False)
        pivot_table_data_status_bpkb.to_excel(writer, sheet_name='Hasil Pivot data Status BPKB')
        pivot_table_status_data_kendaraan.to_excel(writer, sheet_name='Hasil Pivot data Kendaraan')
    print("File Excel berhasil dibuat dengan tiga sheet.")



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