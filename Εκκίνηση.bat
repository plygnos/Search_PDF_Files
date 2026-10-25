@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")
%PY% --version >nul 2>nul
if errorlevel 1 (
 echo Χρειάζεται Python 3.10 ή νεότερη. Κατέβασε από https://www.python.org/downloads/windows/
 pause
 exit /b 1
)
if not exist ".venv\Scripts\pythonw.exe" (
 echo Προετοιμασία εφαρμογής...
 %PY% -m venv .venv
 if errorlevel 1 (pause & exit /b 1)
 ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
 if errorlevel 1 (pause & exit /b 1)
)
start "" ".venv\Scripts\pythonw.exe" "search_files.py"
