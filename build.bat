@echo off
cd /d "%~dp0"
pyinstaller --onefile --windowed --name 입출금정리 --collect-all winsdk --collect-all winocr main.py
echo.
echo 빌드 완료: dist\입출금정리.exe
pause
