@echo off
cls

CALL conda create --name minipayrollsystem python=3.9.5 -y
CALL conda activate minipayrollsystem
CALL yes | pip install openpyxl streamlit pandas datetime

echo [44mMini Payroll System Berhasil di Konfigurasi[0m
echo [44mDouble-click start_mps.bat to start the app...[0m

set/p<nul =Any key to exit ...&pause>nul