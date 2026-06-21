@echo off
REM Versao silenciosa para o Agendador do Windows (nao mexe na planilha
REM se ela estiver aberta - gera so dashboard + snapshot; comente --no-xlsx
REM se quiser que tambem atualize o .xlsx automaticamente).
cd /d "%~dp0"
python -m pip install --quiet requests openpyxl 2>nul
python kalshi_update.py --no-xlsx >> kalshi_log.txt 2>&1
python salvar_snapshot.py >> kalshi_log.txt 2>&1
