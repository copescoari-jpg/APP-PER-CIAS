@echo off
cd /d "%~dp0"
python3.12 -m streamlit run main.py --server.headless false --browser.gatherUsageStats false
if errorlevel 1 (
    echo.
    echo Erro ao iniciar. Pressione qualquer tecla...
    pause > nul
)
