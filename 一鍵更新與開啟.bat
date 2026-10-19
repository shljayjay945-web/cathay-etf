@echo off
title Cathay Fund Update and Sync
cd /d "%~dp0"
python update_and_sync.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ======================================================================
    echo  Execution finished with error. Please see messages above.
    echo ======================================================================
    pause
) else (
    echo.
    echo ======================================================================
    echo  Done! Window will close in 5 seconds.
    echo ======================================================================
    ping 127.0.0.1 -n 6 > nul
)
