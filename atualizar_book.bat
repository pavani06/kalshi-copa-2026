@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Sincronizando o book a partir da boletagem...
python atualizar_book.py %*
if errorlevel 1 (
  echo.
  echo [ERRO] Falhou. Verifique se o boletagem_export.xlsx e os precos estao atualizados.
)
echo.
pause
