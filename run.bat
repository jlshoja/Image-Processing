@echo off
cd /d "%~dp0"

:menu
cls
echo ===============================
echo   Image Processing
echo ===============================
echo   1. Convert images to WebP only
echo   2. Process images (center, white bg, color naming, size limits from settings.json)
echo   3. Full pipeline (convert + process)
echo   4. Exit
echo ===============================
choice /c 1234 /n /m "Choose option: "

if errorlevel 4 goto end
if errorlevel 3 goto full
if errorlevel 2 goto process
if errorlevel 1 goto convert

:convert
python process_images.py convert
echo.
pause
goto menu

:process
python process_images.py process
echo.
pause
goto menu

:full
python process_images.py full
echo.
pause
goto menu

:end
