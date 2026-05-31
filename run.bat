@echo off
REM Avvia easy_marc.py nel venv locale.
REM Uso: run.bat <file.iso> [--config config.json] [--output out.xlsx]

SET SCRIPT_DIR=%~dp0
SET VENV=%SCRIPT_DIR%.venv

IF NOT EXIST "%VENV%" (
    echo Creo l'ambiente virtuale...
    python -m venv "%VENV%"
    "%VENV%\Scripts\pip" install --quiet -r "%SCRIPT_DIR%requirements.txt"
    echo Ambiente pronto.
)

"%VENV%\Scripts\python" "%SCRIPT_DIR%easy_marc.py" %*
