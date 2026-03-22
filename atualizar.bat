@echo off
cd /d "%~dp0"
taskkill /f /im excel.exe >nul 2>&1
timeout /t 2 /nobreak >nul
git fetch origin
git stash push -u
git reset --hard origin/main
git pull origin main
echo.
echo ATUALIZADO com sucesso! Pode abrir o Excel novamente.
pause