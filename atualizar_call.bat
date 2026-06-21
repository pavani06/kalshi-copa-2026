@echo off
REM ===================================================================
REM  atualizar_call.bat - le os precos do Call da Copa que voce colou
REM  em call_da_copa.txt (copiado do WhatsApp) e atualiza a planilha
REM  e o dashboard.
REM  Passo 1: cole as mensagens do grupo em call_da_copa.txt e salve.
REM  Passo 2: (planilha fechada) clique duplo aqui.
REM ===================================================================
cd /d "%~dp0"
echo.
echo == Lendo precos do Call da Copa (call_da_copa.txt) ==
echo.
python -m pip install --quiet openpyxl 2>nul
python ler_call_da_copa.py
echo.
echo == Refletindo no dashboard ==
python kalshi_update.py --no-xlsx
python salvar_snapshot.py
echo.
echo Pronto. Confira a coluna Call da Copa na planilha e no dashboard.
pause
