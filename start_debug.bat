@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo  문제 확인용 실행입니다. 이 창의 내용과 data\logs\server.log 를 확인하세요.
echo.
if exist "python\python.exe" (
  "python\python.exe" "app\server.py"
  goto done
)
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "app\server.py"
  goto done
)
where python >nul 2>nul
if %errorlevel%==0 (
  python "app\server.py"
  goto done
)
echo [오류] 사용할 수 있는 Python을 찾지 못했습니다.
echo setup_python.bat 을 먼저 실행하세요.
:done
echo.
pause
