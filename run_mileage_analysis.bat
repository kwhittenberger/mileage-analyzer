@echo off
REM Mileage Analysis Tool for D'Ewart Representatives, L.L.C.
REM This script analyzes trip data and generates a mileage report

echo ================================================================================
echo MILEAGE ANALYSIS TOOL
echo D'Ewart Representatives, L.L.C.
echo ================================================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Check if CSV file exists
if not exist "volvo-trips-log.csv" (
    echo ERROR: volvo-trips-log.csv not found in current directory
    echo Please place your CSV file in the same folder as this script
    pause
    exit /b 1
)

echo Analyzing mileage data...
echo.

REM Check if user wants business lookup
set /p LOOKUP="Enable business lookup? (slower but more accurate) [y/N]: "

if /i "%LOOKUP%"=="y" (
    echo Running analysis with business lookup...
    python analyze_mileage.py volvo-trips-log.csv --lookup > mileage_report.txt 2>&1
) else (
    echo Running analysis...
    python analyze_mileage.py volvo-trips-log.csv > mileage_report.txt 2>&1
)

if %errorlevel% equ 0 (
    echo.
    echo ================================================================================
    echo Analysis complete!
    echo.
    echo Files generated:
    echo   - mileage_report.txt (text report)
    echo   - weekly_summary.csv (Excel-compatible)
    echo   - detailed_trips.csv (Excel-compatible)
    echo   - summary.csv (Excel-compatible)
    echo ================================================================================
    echo.
    echo Opening text report in Notepad...
    notepad mileage_report.txt
) else (
    echo.
    echo ERROR: Analysis failed. Check mileage_report.txt for details
    pause
    exit /b 1
)

pause
