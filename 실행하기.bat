@echo off
chcp 65001 > nul
echo ===================================================
echo   PDF Studio PRO - 실행 중...
echo ===================================================
echo.
echo 브라우저에서 PDF Studio를 엽니다.
start "" "%~dp0index.html"
exit
