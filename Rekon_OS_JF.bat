@echo off
REM Skrip ini akan menjalankan Python script di folder yang sama dengan file .bat ini

REM %~dp0 merujuk ke folder tempat file batch ini berada
set SCRIPT_PATH=%~dp0Recon_JF_OS_Gabungan.py

REM Menjalankan script menggunakan perintah python sistem (pastikan Python sudah masuk PATH)
python "%SCRIPT_PATH%"
pause
