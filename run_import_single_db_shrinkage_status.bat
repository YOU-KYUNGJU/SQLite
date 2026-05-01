@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"
set YEAR=%~1
if "%YEAR%"=="" set YEAR=%date:~0,4%

call :find_python
if not defined PYTHON_CMD (
    echo.
    echo [ERROR] No usable Python was found.
    echo [ERROR] Install real Python or set PYTHON_EXE.
    pause
    exit /b 1
)

%PYTHON_CMD% src\run_part_import.py --config config\parts\shrinkage.json --year %YEAR% --replace-db

if errorlevel 1 (
    echo.
    echo Import failed. Check the logs folder for details.
    pause
    exit /b 1
)

echo.
echo Import completed. Check the logs folder for details.
pause
exit /b 0

:find_python
set "PYTHON_CMD="

if defined PYTHON_EXE (
    if exist "%PYTHON_EXE%" (
        set "PYTHON_CMD="%PYTHON_EXE%""
        goto :eof
    )
)

for %%P in (
    "%LocalAppData%\Programs\Python\Python314\python.exe"
    "%LocalAppData%\Programs\Python\Python313\python.exe"
    "%LocalAppData%\Programs\Python\Python312\python.exe"
    "%LocalAppData%\Programs\Python\Python311\python.exe"
    "%ProgramFiles%\Python314\python.exe"
    "%ProgramFiles%\Python313\python.exe"
    "%ProgramFiles%\Python312\python.exe"
    "%ProgramFiles%\Python311\python.exe"
) do (
    if exist %%~P (
        set "PYTHON_CMD=%%~fP"
        goto :eof
    )
)

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :eof
)

where python >nul 2>nul
if %errorlevel%==0 (
    python -c "import sys; print(sys.executable)" >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
        goto :eof
    )
)

goto :eof
