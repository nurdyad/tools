@echo off
echo 🧹 Cleaning old build data...
if exist build rd /s /q build
if exist dist rd /s /q dist

echo 🔨 Building EXE...
:: This line now uses the correct filename from your screenshot
pyinstaller --onefile --noconfirm --add-data "check-ods-mismatch.ps1;." onboarding.py

echo.
if %errorlevel% equ 0 (
    echo ✅ Success! Your tool is in the 'dist' folder.
) else (
    echo ❌ Build failed.
)
pause