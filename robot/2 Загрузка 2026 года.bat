@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Загрузка идёт несколько часов. Компьютер не выключайте; если прервётся - запустите снова, скачанное не повторится.
"%~dp0.venv\Scripts\python.exe" run.py load 2026
echo.
echo Готово. Excel для оператора - в папке data\out
pause
