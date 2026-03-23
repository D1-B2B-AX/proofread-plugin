@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo.
echo ========================================
echo   교안 검수 플러그인 설치
echo ========================================
echo.

:: 현재 스크립트 위치 = 플러그인 디렉토리
set "PLUGIN_DIR=%~dp0"
:: 끝의 \ 제거
set "PLUGIN_DIR=%PLUGIN_DIR:~0,-1%"
:: 백슬래시를 슬래시로 변환
set "PLUGIN_DIR=%PLUGIN_DIR:\=/%"

echo [1/4] 플러그인 경로: %PLUGIN_DIR%

:: Python 확인
echo.
echo [2/4] Python 확인 중...
python3 --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    python --version >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo [오류] Python이 설치되어 있지 않습니다.
        echo Python 3.10 이상을 설치해 주세요.
        pause
        exit /b 1
    )
)
echo Python 확인 완료

:: 필요 패키지 설치
echo.
echo [3/4] 필요 패키지 설치 중...
pip install python-pptx PyPDF2 2>nul || pip3 install python-pptx PyPDF2 2>nul
echo 패키지 설치 완료

:: 스킬 파일 복사 (경로 치환)
echo.
echo [4/4] 스킬 파일 설치 중...

set "COMMANDS_DIR=%USERPROFILE%\.claude\commands"
if not exist "%COMMANDS_DIR%" mkdir "%COMMANDS_DIR%"

:: 교안검수.md에서 {{PLUGIN_DIR}}을 실제 경로로 치환하여 복사
set "SOURCE=%~dp0교안검수.md"
set "TARGET=%COMMANDS_DIR%\교안검수.md"

powershell -Command "(Get-Content '%SOURCE%' -Raw -Encoding UTF8) -replace '\{\{PLUGIN_DIR\}\}', '%PLUGIN_DIR%' | Set-Content '%TARGET%' -Encoding UTF8"

echo 스킬 파일 설치 완료: %TARGET%

:: claude_temp 폴더 생성
if not exist "%USERPROFILE%\Desktop\claude_temp" mkdir "%USERPROFILE%\Desktop\claude_temp"

echo.
echo ========================================
echo   설치 완료!
echo ========================================
echo.
echo 사용법:
echo   1. Claude Code 실행
echo   2. /교안검수 입력
echo   3. 교안 파일을 드래그하거나 경로 입력
echo.
echo ========================================
pause
