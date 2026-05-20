@echo off
cd /d "%~dp0"

echo ================================================
echo   ComplyAdvantage — Actualisation des donnees
echo ================================================
echo.

REM Demander la periode
set /p DAYS="Periode (nb de jours, defaut 30) : "
if "%DAYS%"=="" set DAYS=30

echo.
echo Recuperation des donnees (%DAYS% jours)...
echo.

python fetch_data.py --days %DAYS%

if errorlevel 1 (
    echo.
    echo ERREUR lors de la recuperation des donnees.
    echo Verifiez que Python est installe et que la cle API est correcte.
    pause
    exit /b 1
)

echo.
echo Deploiement sur GitHub...
echo.

cd /d "C:\Users\ThomasOliveira\OneDrive - EIFFEL IG\Documents\Github\compliance-app"

git add complyadvantage/data.json
git commit -m "data: refresh ComplyAdvantage (%DAYS%j)"
git push

echo.
echo ================================================
echo   Done ! Le dashboard est a jour.
echo ================================================
echo.
pause
