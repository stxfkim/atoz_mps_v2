@echo off
cls
CALL conda activate minipayrollsystem
CALL streamlit run main.py

set/p<nul =Any key to exit ...&pause>nul
