@echo off
REM ===================================================================
REM  atualizar_kalshi.bat - puxa dados AO VIVO da Kalshi e atualiza
REM  a planilha do modelo do pico (Copa 2026).
REM  IMPORTANTE: feche a planilha no Excel antes de rodar.
REM ===================================================================
cd /d "%~dp0"
echo.
echo == Atualizando odds da Kalshi (KXMENWORLDCUP) ==
echo.
python -m pip install --quiet requests openpyxl 2>nul
python kalshi_update.py
python salvar_snapshot.py
echo.
echo Pronto. Abra kalshi_dashboard.html para ver, ou a planilha para operar.
pause
