@echo off
setlocal
cd /d "%~dp0"

echo [1/4] Checking for 64-bit Python...
python -c "import struct,sys; sys.exit(0 if struct.calcsize('P')*8==64 else 1)"
if errorlevel 1 goto :no64

echo [2/4] Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :fail

echo [3/4] Building the portable exe (PyInstaller)...
python -m PyInstaller --noconfirm --clean BatchCsvConverter.spec
if errorlevel 1 goto :fail
if not exist release mkdir release
copy /y dist\BatchCsvConverter.exe release\BatchCsvConverter_Portable_x64.exe >nul

echo [4/4] Building the installer (Inno Setup)...
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC goto :noinno
"%ISCC%" setup_script.iss
if errorlevel 1 goto :fail
goto :done

:noinno
echo Inno Setup 6 was not found - skipping the installer.
echo Open setup_script.iss manually in Inno Setup and press F9.
goto :done

:no64
echo ERROR: 64-bit Python is required for an x64 build.
pause
exit /b 1

:fail
echo.
echo BUILD FAILED - see the messages above.
pause
exit /b 1

:done
echo.
echo ============================================================
echo  Done. Files are in the "release" folder:
echo    BatchCsvConverter_Portable_x64.exe   (portable, no install)
echo    BatchCsvConverter_Setup_x64.exe      (installer)
echo ============================================================
pause
