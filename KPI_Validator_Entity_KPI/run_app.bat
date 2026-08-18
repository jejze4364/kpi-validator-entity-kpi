@echo off
cd /d "%~dp0"
if not exist .venv (
 py -m venv .venv
 call .venv\Scripts\activate
 python -m pip install --upgrade pip
 pip install -r requirements.txt
) else (
 call .venv\Scripts\activate
)
start "" http://localhost:8501
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
