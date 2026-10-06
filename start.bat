@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist "app\server.py" goto broken
if not exist "app\web\index.html" goto broken

rem --- 폴더 안 포터블 파이썬 우선 ---
if exist "python\pythonw.exe" (
  start "" "python\pythonw.exe" "app\server.py"
  exit /b 0
)

rem --- PC에 설치된 파이썬 사용 ---
where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw -3 "app\server.py"
  exit /b 0
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
  start "" pythonw "app\server.py"
  exit /b 0
)

rem pythonw가 PATH에 없더라도 일반 python/py가 있으면 최소화 창으로 실행
where py >nul 2>nul
if %errorlevel%==0 (
  start "제천내신전보" /min py -3 "app\server.py"
  exit /b 0
)
where python >nul 2>nul
if %errorlevel%==0 (
  start "제천내신전보" /min python "app\server.py"
  exit /b 0
)

echo.
echo  [안내] 파이썬이 준비되지 않았습니다.
echo  같은 폴더의 setup_python.bat 을 먼저 실행하세요.
echo.
pause
exit /b 1

:broken
echo.
echo  [오류] 프로그램 구성 파일이 없습니다.
echo  app\server.py 와 app\web\index.html 이 있는지 확인하세요.
echo  압축파일을 폴더째 다시 풀어 주세요.
echo.
pause
exit /b 2
