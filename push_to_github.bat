@echo off
title Push to GitHub - RMA Schematic Finder

echo.
echo ============================================
echo   Push to NEW repo: RMA-Schematic-Finder
echo ============================================
echo.

cd /d "%~dp0"

:: Remove old remote if exists
git remote remove origin >nul 2>&1

:: Init git if needed
if not exist ".git" (
    echo [1/4] Initializing git...
    git init
) else (
    echo [1/4] Git already initialized.
)

echo [2/4] Adding files...
git add .

echo [3/4] Creating commit...
git commit -m "feat: MOXA Schematic Link Finder v1.2"
if errorlevel 1 (
    echo       No new changes to commit, will push existing commits.
)

echo [4/4] Creating NEW GitHub repo: RMA-Schematic-Finder ...
gh repo create RMA-Schematic-Finder --public --source=. --remote=origin --push

if errorlevel 1 goto :err

echo.
echo ============================================
echo   Done! New repo created:
echo   RMA-Schematic-Finder
echo ============================================
echo.
pause
exit /b 0

:err
echo.
echo [ERROR] Something went wrong. Please check the output above.
pause
exit /b 1
