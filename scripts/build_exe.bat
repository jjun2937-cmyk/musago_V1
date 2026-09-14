@echo off
REM Windows에서 실행: 무사고전환 보고서 생성기 exe 빌드 스크립트.
REM 사전 준비: 파이썬 3.10+ 이 설치되어 있고 PATH에 python 명령이 잡혀 있어야 함.
setlocal

cd /d "%~dp0\.."

echo [1/3] 의존성 설치 중...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/3] 이전 빌드 정리 중...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] exe 빌드 중(PyInstaller)...
python -m PyInstaller musago_report.spec
if errorlevel 1 goto :error

echo.
echo 빌드 완료: dist\무사고전환_보고서_생성기.exe
pause
exit /b 0

:error
echo.
echo 빌드 중 오류가 발생했습니다. 위 로그를 확인해 주세요.
pause
exit /b 1
