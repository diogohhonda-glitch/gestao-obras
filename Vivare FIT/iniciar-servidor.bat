@echo off
cd /d "%~dp0"
echo Iniciando servidor local Vivare FIT...
echo Acesse: http://localhost:8000
start http://localhost:8000
python -m http.server 8000
pause
