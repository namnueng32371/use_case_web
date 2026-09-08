@echo off
echo Starting MySQL...
start "MySQL" /min "C:\Program Files\MySQL\MySQL Server 8.4\bin\mysqld.exe" --defaults-file="C:\ProgramData\MySQL\MySQL Server 8.4\my.ini" --console
timeout /t 4 /nobreak >nul

echo Starting AI Use Case app server...
cd /d "%~dp0.."
python server.py

pause
