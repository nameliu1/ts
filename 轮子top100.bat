@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Starting the process...

if exist port.txt del /F /Q port.txt
if exist url.txt del /F /Q url.txt

python 2.py
python ppp.py

if exist res.json del /F /Q res.json
if exist res_processed.txt del /F /Q res_processed.txt
if exist res_processed.xlsx del /F /Q res_processed.xlsx

start python 1.py

timeout /t 60 /nobreak >nul

echo Done
pause