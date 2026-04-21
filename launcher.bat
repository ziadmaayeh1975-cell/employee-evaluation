@echo off
title Fannoun Performance System

if not exist "app.py" (
    echo ERROR: app.py not found
    pause
    exit /b 1
)

if not exist "final Apprisal.xlsm" (
    echo ERROR: final Apprisal.xlsm not found
    pause
    exit /b 1
)

set LOCAL_IP=
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4" ^| findstr /v "169.254"') do (
    if not defined LOCAL_IP set LOCAL_IP=%%a
)
set LOCAL_IP=%LOCAL_IP: =%

cls
echo.
echo  +--------------------------------------------------+
echo  ^|   Fannoun Performance Evaluation System          ^|
echo  +--------------------------------------------------+
echo  ^|                                                  ^|
echo  ^|   Local:    http://localhost:8501                ^|
echo  ^|   Network:  http://%LOCAL_IP%:8501
echo  ^|                                                  ^|
echo  ^|   Share network link with other users            ^|
echo  ^|   Close this window to stop the system           ^|
echo  ^|                                                  ^|
echo  +--------------------------------------------------+
echo.

timeout /t 3 /nobreak >nul
start http://localhost:8501

python -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true --browser.gatherUsageStats=false --server.enableCORS=false --server.enableXsrfProtection=false

echo.
echo  System stopped. Press any key to exit.
pause >nul
