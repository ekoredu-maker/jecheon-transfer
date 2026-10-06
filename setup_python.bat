@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PYVER=3.12.10"
set "ZIP=python-%PYVER%-embed-amd64.zip"
set "URL=https://www.python.org/ftp/python/%PYVER%/%ZIP%"

if exist "python\python.exe" goto already

if exist "%ZIP%" goto unzip

echo.
echo  공식 사이트 python.org 에서 파이썬 %PYVER% 을 내려받습니다. 약 10MB
echo  %URL%
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri '%URL%' -OutFile '%ZIP%'"
if not exist "%ZIP%" goto nodownload

:unzip
echo  압축을 푸는 중...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Force -LiteralPath '%ZIP%' -DestinationPath 'python'"
if not exist "python\python.exe" goto fail
del /q "%ZIP%" >nul 2>nul
echo.
echo  완료되었습니다. 이제 start.bat 으로 실행하세요.
echo.
pause
exit /b 0

:already
echo.
echo  파이썬이 이미 준비되어 있습니다. start.bat 으로 실행하세요.
echo.
pause
exit /b 0

:nodownload
echo.
echo  [내려받기 실패] 인터넷이 안 되는 PC일 수 있습니다.
echo  인터넷이 되는 PC에서 아래 주소의 파일을 받아
echo  %URL%
echo  이 폴더 %~dp0 에 그대로 넣은 뒤 setup_python.bat 을 다시 실행하세요.
echo.
pause
exit /b 1

:fail
echo.
echo  [오류] 압축 풀기에 실패했습니다. %ZIP% 파일이 손상되었을 수 있으니 지우고 다시 실행하세요.
echo.
pause
exit /b 1
