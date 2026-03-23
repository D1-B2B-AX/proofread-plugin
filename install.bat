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
set "PLUGIN_DIR=%PLUGIN_DIR:~0,-1%"
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
if %ERRORLEVEL% NEQ 0 (
    echo [경고] 패키지 설치에 문제가 있을 수 있습니다.
    echo 이미 설치되어 있다면 무시해도 됩니다.
)
echo 패키지 설치 완료

:: 스킬 파일 복사 (cmd 네이티브 방식 — powershell 미사용)
echo.
echo [4/4] 스킬 파일 설치 중...

set "COMMANDS_DIR=%USERPROFILE%\.claude\commands"
if not exist "%COMMANDS_DIR%" (
    mkdir "%COMMANDS_DIR%"
    if %ERRORLEVEL% NEQ 0 (
        echo [오류] commands 폴더 생성에 실패했습니다.
        echo 경로: %COMMANDS_DIR%
        pause
        exit /b 1
    )
)

set "SOURCE=%~dp0교안검수.md"
set "TARGET=%COMMANDS_DIR%\교안검수.md"

:: 원본 파일 존재 확인
if not exist "%SOURCE%" (
    echo [오류] 교안검수.md 파일을 찾을 수 없습니다.
    echo 경로: %SOURCE%
    echo 다운로드한 폴더 안에 교안검수.md가 있는지 확인해 주세요.
    pause
    exit /b 1
)

:: cmd 네이티브 방식으로 {{PLUGIN_DIR}} 치환
set "SEARCH={{PLUGIN_DIR}}"
set "REPLACE=%PLUGIN_DIR%"

> "%TARGET%" (
    for /f "usebackq tokens=* delims=" %%a in ("%SOURCE%") do (
        set "line=%%a"
        if defined line (
            echo !line:%SEARCH%=%REPLACE%!
        ) else (
            echo.
        )
    )
)

:: 복사 결과 확인
if not exist "%TARGET%" (
    echo [오류] 스킬 파일 복사에 실패했습니다.
    echo.
    echo 해결 방법:
    echo   1. 다운로드 경로에 한글이 포함되어 있다면
    echo      C:\proofread-plugin 등 영문 경로로 옮긴 후 다시 시도해 주세요.
    echo   2. 또는 교안검수.md 파일을 직접 아래 폴더에 복사해 주세요:
    echo      %COMMANDS_DIR%
    pause
    exit /b 1
)
echo 스킬 파일 설치 완료: %TARGET%

:: claude_temp 폴더 생성
if not exist "%USERPROFILE%\Desktop\claude_temp" (
    mkdir "%USERPROFILE%\Desktop\claude_temp"
    if %ERRORLEVEL% NEQ 0 (
        echo [경고] claude_temp 폴더 생성에 실패했습니다.
        echo 바탕화면에 claude_temp 폴더를 직접 만들어 주세요.
    ) else (
        echo claude_temp 폴더 생성 완료
    )
) else (
    echo claude_temp 폴더 이미 존재
)

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
