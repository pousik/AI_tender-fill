@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  py -3.13 -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install -U pip
python -m pip install -r requirements.txt
python -m models.fill_data_db
python tools/add_knowledge.py knowledge_seed.json
if "%~1"=="" (
  echo.
  echo Использование: run.bat "input.docx" "output.docx"
  echo Сначала задайте GIGACHAT_CREDENTIALS в окружении.
  exit /b 1
)
if "%~2"=="" set "OUT=Заполненный_%~nx1" else set "OUT=%~2"
python main.py "%~1" "%OUT%"
