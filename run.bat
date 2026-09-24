@echo off
echo ========================================================
echo   Menjalankan Generator Laporan Naratif Dinkes...
echo ========================================================
uv run streamlit run app.py --server.port 8501 --server.address 127.0.0.1
pause
