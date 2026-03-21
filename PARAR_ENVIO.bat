@echo off
REM Fechar o arquivo específico Envios_WhatsApp.xlsm (só se for o único Excel aberto)
taskkill /f /im "excel.exe"
echo Excel foi fechado.
pause
