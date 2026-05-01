@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"
set CONFIG=%~1
set YEAR=%~2

if "%CONFIG%"=="" (
    echo [ERROR] config 파일 경로가 필요합니다.
    echo 예시: run_import_with_config.bat config\parts\shrinkage.json 2026
    pause
    exit /b 1
)

if "%YEAR%"=="" set YEAR=%date:~0,4%

call :find_python
if not defined PYTHON_CMD (
    echo [ERROR] No usable Python was found.
    pause
    exit /b 1
)

%PYTHON_CMD% src\run_part_import.py --config "%CONFIG%" --year %YEAR% --replace-db
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
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :eof
)
where python >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :eof
)
goto :eof
