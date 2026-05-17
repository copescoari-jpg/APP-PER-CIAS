@echo off
cd /d "%~dp0"
python3.12 app.py
if errorlevel 1 (
    echo.
    echo Erro ao iniciar. Pressione qualquer tecla...
    pause > nul
)
