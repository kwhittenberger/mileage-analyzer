@echo off
REM Build the Mileage Analyzer GUI executable
REM D'Ewart Representatives, L.L.C.

echo ============================================================
echo Building Mileage Analyzer GUI Executable
echo ============================================================
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building executable from spec file...
echo This may take several minutes due to PyQt6 WebEngine components.
echo.

pyinstaller MileageAnalyzerGUI.spec --noconfirm

echo.
if exist "dist\MileageAnalyzerGUI.exe" (
    echo ============================================================
    echo BUILD SUCCESSFUL!
    echo ============================================================
    echo.
    echo Executable created at: dist\MileageAnalyzerGUI.exe
    echo.
    echo To run, double-click MileageAnalyzerGUI.exe in the dist folder.
) else (
    echo ============================================================
    echo BUILD FAILED
    echo ============================================================
    echo Please check the error messages above.
)

echo.
pause
