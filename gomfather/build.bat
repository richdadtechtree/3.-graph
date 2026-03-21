@echo off
chcp 65001 > nul
echo ========================================
echo   gomfather-kb EXE 빌드 시작
echo ========================================

cd /d "%~dp0"

echo.
echo [1/2] 기존 빌드 폴더 정리 중...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [2/2] PyInstaller 빌드 중...
pyinstaller gomfather-kb.spec

echo.
if exist dist\gomfather-kb.exe (
    echo ========================================
    echo   빌드 성공!
    echo   출력 파일: dist\gomfather-kb.exe
    echo ========================================
    explorer dist
) else (
    echo ========================================
    echo   빌드 실패! 위 오류 메시지를 확인하세요.
    echo ========================================
)

pause
