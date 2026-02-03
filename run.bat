@echo off
echo ================================
echo  INICIANDO ENVIO DE CORREOS
echo ================================

cd /d "%~dp0"

REM Ejecutar directamente el python del entorno virtual
.\venv\Scripts\python.exe send-mail.py

echo ================================
echo  PROCESO FINALIZADO
echo ================================
pause
