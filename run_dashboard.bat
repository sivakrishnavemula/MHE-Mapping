@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Could not install the dashboard requirements. Check that Python and pip are installed.
    pause
    exit /b 1
)
python -m streamlit run app.py
if errorlevel 1 (
    echo.
    echo The dashboard could not be started. Check the message above.
    pause
)
