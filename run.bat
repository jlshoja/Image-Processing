@echo off
cd /d "%~dp0"

:menu
cls
echo ===============================
echo   Image Processing
echo ===============================
echo   1. Convert + compress to WebP under limit (keeps original names)
echo   2. Rename photoroom images by color (to "final image renamed")
echo   3. Exit
echo ===============================
choice /c 123 /n /m "Choose option: "

if errorlevel 3 goto end
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

:end
