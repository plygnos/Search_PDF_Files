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
if not exist ".venv\Scripts\python.exe" (
 %PY% -m venv .venv
 if errorlevel 1 (pause & exit /b 1)
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check --upgrade -r requirements.txt pyinstaller
if errorlevel 1 (pause & exit /b 1)
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name "Anazitisi_Arxeion" --collect-all fitz --collect-all pymupdf --collect-all striprtf "search_files.py"
if errorlevel 1 (pause & exit /b 1)
echo.
echo Το εκτελέσιμο βρίσκεται στον φάκελο dist\Anazitisi_Arxeion.exe
pause
