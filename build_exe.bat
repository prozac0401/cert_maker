@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

echo [1/4] Python 가상환경 준비...
if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv
  if errorlevel 1 goto :error
)

echo [2/4] 빌드 의존성 설치...
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [3/4] Windows 실행파일 빌드...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name CertificateMaker ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32timezone ^
  --hidden-import win32com.client ^
  --hidden-import win32com.client.dynamic ^
  --hidden-import win32com.client.gencache ^
  --collect-submodules win32com ^
  --add-data "samples;samples" ^
  certificate_maker.py
if errorlevel 1 goto :error

echo [4/4] 빌드 완료
echo EXE: %cd%\dist\CertificateMaker.exe
pause
exit /b 0

:error
echo.
echo 빌드 실패. Python 및 pywin32 설치 상태를 확인하세요.
pause
exit /b 1
