@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not exist ".venv\Scripts\python.exe" (
  echo .venv가 없습니다. 먼저 build_exe.bat를 한 번 실행해 환경을 구성하세요.
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"
python certificate_maker.py
