@echo off
cd /d "%~dp0"
echo Starting TMB Payroll app...
echo Once it says "Running on http://127.0.0.1:5000", open that address in your browser.
echo Keep this window open while you use the app - closing it stops the server.
echo.
python app.py
pause
