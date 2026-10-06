@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo  제천 내신전보 포터블 구성 진단을 실행합니다.
echo.
if exist "python\python.exe" (
  "python\python.exe" "app\server.py" --self-test
  goto done
)
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "app\server.py" --self-test
  goto done
)
where python >nul 2>nul
if %errorlevel%==0 (
  python "app\server.py" --self-test
  goto done
)
echo [오류] Python을 찾지 못했습니다. setup_python.bat 을 먼저 실행하세요.
:done
echo.
pause
