@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo Installation des corrections VBA...
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass ^
  -File "%~dp0scripts\Sync-VbaProject.ps1"

if errorlevel 1 (
  echo.
  echo ECHEC : lisez le message ci-dessus.
  echo Le classeur modele Projet_VBA_BS.xlsm est reste intact.
  pause
  exit /b 1
)

echo.
echo TERMINE : ouvrez maintenant Projet_VBA_BS_corrige.xlsm
pause
