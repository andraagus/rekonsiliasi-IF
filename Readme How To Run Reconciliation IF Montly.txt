**Panduan Penggunaan Alat Rekonsiliasi IF Bulanan**

Dokumen ini menjelaskan langkah-langkah yang diperlukan untuk menjalankan alat rekonsiliasi bulanan untuk data BPKB dan Outstanding (OS) dengan berbagai partner.

---

### 1. Prasyarat (Requirement)

Sebelum menjalankan alat rekonsiliasi, pastikan Anda telah memenuhi prasyarat berikut:

*   **Instalasi Python:** Pastikan Python sudah terinstal di sistem Anda. Anda bisa mengunduhnya dari [python.org](https://www.python.org/).
*   **Instalasi Library Python:** Instal library yang dibutuhkan menggunakan pip:
    ```bash
    pip install pandas numpy tqdm openpyxl
    ```

---

### 2. Sumber Data (Data Source)

Semua file data yang akan digunakan untuk proses rekonsiliasi **harus dikumpulkan dalam satu folder yang sama** dengan skrip Python (`.py`) dan file batch (`.bat`).

**A. Data Internal (Data Akhir Bulan - EOM)**

Data ini berasal dari sistem internal dan harus disiapkan setiap akhir bulan (EOM).

1.  `ASSETINVENTORY.txt`: Data inventaris aset (asset inventory) akhir bulan.
2.  `WRITEOFF.txt`: Data aset yang dihapus buku (write-off) akhir bulan.
3.  `CORACCOUNT_BJJ.xlsx`: Data akun inti (core account) BJJ akhir bulan, yang telah dikonversi ke format Excel.
4.  `OBJECTCAR_BJJ.xlsx`: Data objek kendaraan (object car) BJJ akhir bulan, yang telah dikonversi ke format Excel.

**B. Data Partner (Dikonversi Sesuai Template)**

Data ini berasal dari masing-masing partner dan harus dikonversi ke template yang sesuai sebelum digunakan.

*   **Partner ACC (Gabungan ASF dan SBSF):**
    5.  `BPKB_ACC.xlsx`: Data BPKB dari ACC.
    6.  `OS_ACC.xlsx`: Data Outstanding (OS) dari ACC.

*   **Partner FIF:**
    7.  `BPKB_FIF.xlsx`: Data BPKB dari FIF.
    8.  `OS_FIF.xlsx`: Data Outstanding (OS) dari FIF.

*   **Partner TAF:**
    9.  `BPKB_TAF.xlsx`: Data BPKB dari TAF.
    10. `OS_TAF.xlsx`: Data Outstanding (OS) dari TAF.

*   **Partner Amartha:**
    11. `OS_AMARTHA.xlsx`: Data Outstanding (OS) dari Amartha.

*   **Partner Easycash:**
    12. `OS_EASYCASH.xlsx`: Data Outstanding (OS) dari Easycash.

*   **Partner Finture:**
    13. `OS_FINTURE.xlsx`: Data Outstanding (OS) dari Finture.

*   **Partner HCI:**
    14. `OS_HCI.xlsx`: Data Outstanding (OS) dari HCI.

---

### 3. Cara Menjalankan Rekonsiliasi

Untuk memulai proses rekonsiliasi, cukup **double-click** pada salah satu file `.bat` berikut:

*   **A. `RekonBPKBJF.bat`**: Digunakan untuk rekonsiliasi data BPKB. Setelah dijalankan, Anda akan diminta untuk memilih partner (ACC, FIF, TAF) atau menjalankan semua rekonsiliasi BPKB.
*   **B. `Rekon_OS_JF.bat`**: Digunakan untuk rekonsiliasi data Outstanding (OS). Setelah dijalankan, Anda akan diminta untuk memilih partner (ACC, FIF, TAF, Amartha, Easycash, Finture, HCI) atau menjalankan semua rekonsiliasi OS.

---

### 4. Source Code Rekonsiliasi

Skrip Python yang digunakan untuk proses rekonsiliasi adalah:

*   `Recon_JF_BPKB_Gabungan.py`
*   `Recon_JF_OS_Gabungan.py`

---

### 5. Analisis Hasil Rekonsiliasi

Setelah proses rekonsiliasi selesai, file Excel hasil rekonsiliasi akan dibuat di dalam folder `output_bpkb` atau `output_os` (tergantung jenis rekonsiliasi) di direktori yang sama dengan skrip. Hasil ini perlu dianalisis dengan cermat:

*   **Untuk Data Outstanding (OS) yang Berselisih:**
    *   Data yang menunjukkan selisih OS harus dikonfirmasi terlebih dahulu dengan data pembayaran yang dilakukan pada tanggal 1 dan 2 hari kerja di bulan selanjutnya. Ini untuk memastikan apakah selisih tersebut disebabkan oleh pembayaran yang belum tercatat pada akhir bulan.
    *   Jika setelah konfirmasi pembayaran masih terdapat selisih, data tersebut harus dikonfirmasi lebih lanjut ke tim bisnis melalui email. Tim bisnis kemudian akan berkoordinasi dengan partner terkait untuk klarifikasi.

*   **Tindak Lanjut (Feedback):**
    *   Setiap umpan balik atau instruksi yang diterima terkait hasil rekonsiliasi (misalnya, koreksi data, investigasi lebih lanjut) harus segera ditindaklanjuti dan diselesaikan.
