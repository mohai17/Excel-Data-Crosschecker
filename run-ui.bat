@echo off
title Excel Comparison Tool

cd /d "%~dp0"

start "" cmd /c python app.py

timeout /t 1 >nul

start http://127.0.0.1:5000

exit