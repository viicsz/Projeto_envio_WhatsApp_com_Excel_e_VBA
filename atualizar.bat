@echo off
cd /d "%~dp0"
git pull origin main
echo.
echo Atualizado! Pode fechar essa janela.
pause