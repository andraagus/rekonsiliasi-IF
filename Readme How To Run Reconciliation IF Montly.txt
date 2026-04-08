-Data Source yang akan digunakan untuk rekon JF
-Semua daya dikumpulkand alam 1 folder 

Data Bank (data akhir bulan)
1. ASSETINVENTORY.txt (data asset inventory EOM)
2. WRITEOFF.txt (data asset WRITEOFF EOM)
3. CORACCOUNT_BJJ.xlsx (data CORACCOUNT_BJJ EOM yang di convert jadi excel terlebihdahulu)
4. OBJECTCAR_BJJ.xlsx (data OBJECTCAR_BJJ_BJJ EOM yang di convert jadi excel terlebihdahulu)

Data ACC (gabungan ASF dan SBSF) dikonversi sesuai template 
5. BPKB_ACC.xlsx
6. OS_ACC.xlsx

Data FIF dikonversi sesuai template 
7. BPKB_FIF.xlsx
8. OS_FIF.xlsx

Data TAF dikonversi sesuai template 
9. BPKB_TAF.xlsx
10.OS_TAF.xlsx

Data Amartha dikonversi sesuai template
11.OS_AMARTHA.xlsx

Data Easycash dikonversi sesuai template
12.OS_EASYCASH.xlsx

Data Finture dikonversi sesuai template
13.OS_FINTURE.xlsx

Data HCI dikonversi sesuai template
14.OS_HCI.xlsx

Run (double klik file ini)
A. RekonBPKBJF.bat (jika akan rekon data BPKB)
B. Rekon_OS_JF.bat (jika akan rekaon data OS)

source code rekonsiliasi
Recon_JF_BPKB_Gabungan.py
Recon_JF_OS_Gabungan.py

Hasil rekon harus di analisa
- untuk data yang OS nya selisih harus dokonfirmasi dengan data pembayaran yang dilakukan tanggal 1 dan 2 hari kerja bulannselanjutnya 
- jika ada data yang masih selisih harus dikonfirmai ke tim bisnis by email agar di konfirmasi ke partner 
- Jika ada feedback yang harus diselesaikan makan segera dijalankan 

