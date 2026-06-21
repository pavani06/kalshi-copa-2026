@echo off
REM ===================================================================
REM  agendar_windows.bat - cria uma tarefa no Agendador do Windows que
REM  roda a atualizacao da Kalshi automaticamente.
REM  Por padrao: todo dia as 08:00. Edite /ST ou /SC abaixo se quiser.
REM  Rode este arquivo UMA vez (clique duplo). Para remover:
REM     schtasks /Delete /TN "Kalshi Copa 2026" /F
REM ===================================================================
cd /d "%~dp0"
schtasks /Create /TN "Kalshi Copa 2026" /TR "\"%~dp0atualizar_kalshi_silent.bat\"" /SC DAILY /ST 08:00 /F
echo.
echo Tarefa "Kalshi Copa 2026" criada (diaria 08:00).
echo Durante os jogos, para rodar de hora em hora, use no Prompt de Comando:
echo    schtasks /Create /TN "Kalshi Copa 2026 horaria" /TR "\"%~dp0atualizar_kalshi_silent.bat\"" /SC HOURLY /F
pause
