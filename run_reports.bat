@echo off
setlocal EnableExtensions DisableDelayedExpansion

if not "%~2"=="" (
    echo Usage: %~nx0 ["path to input folder"]
    exit /b 2
)
set "data_dir=%~f1"
pushd "%~dp0" || exit /b 1
if not defined data_dir set "data_dir=%CD%"
for %%D in ("%data_dir%\.") do set "data_dir=%%~fD"

set "report_python=.venv\Scripts\python.exe"
if not exist "%report_python%" (
    py -3.11 -c "import sys" >nul 2>&1
    if errorlevel 1 (
        python -m venv .venv
    ) else (
        py -3.11 -m venv .venv
    )
    if errorlevel 1 (
        echo Install Python 3.11 and make it available from the command line. 1>&2
        goto :failed
    )
)

set "report_month="
for /f %%M in ('%report_python% -c "from datetime import date, timedelta; previous = date.today().replace(day=1) - timedelta(days=1); print(f'{previous.month:02}.{previous.year:04}')"') do set "report_month=%%M"
if not defined report_month goto :failed

set "input_files="Summary.xlsx" "SKU Transfer Total List.xlsx" "vendorlist.txt" "Calls %report_month%.xlsx" "Emails %report_month%.xlsx" "Gift Cards %report_month%.xlsx" "Shipments %report_month%.xlsx" "Storage %report_month%.xlsx""
for %%F in (%input_files%) do (
    if not exist "%data_dir%\%%~F" (
        echo Missing input file "%data_dir%\%%~F". 1>&2
        goto :failed
    )
)

"%report_python%" -c "from importlib.metadata import version; import sys; sys.exit(version('pandas') != '2.3.2' or version('openpyxl') != '3.1.5')" >nul 2>&1
if errorlevel 1 (
    "%report_python%" -m pip install pandas==2.3.2 openpyxl==3.1.5 || goto :failed
)

if /i not "%data_dir%"=="%CD%" (
    for %%F in (%input_files%) do (
        copy /y "%data_dir%\%%~F" ".\%%~F" >nul || goto :failed
    )
)
if not exist "finish\" mkdir "finish" || goto :failed

"%report_python%" skytransfer.py || goto :failed
"%report_python%" summ.py || goto :failed
"%report_python%" bro.py || goto :failed

echo Updated Summary.xlsx and the vendor workbooks in finish\.
popd
exit /b 0

:failed
echo Report workflow stopped. 1>&2
popd
exit /b 1
