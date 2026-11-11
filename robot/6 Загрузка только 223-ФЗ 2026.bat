@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Загрузка только договоров и протоколов 223-ФЗ за 2026 год (44-ФЗ и 615-ПП пропускаются).
echo Компьютер не выключайте; если прервётся - запустите снова, скачанное не повторится.
"%~dp0.venv\Scripts\python.exe" run.py load 2026 223
echo.
echo Готово. Excel для оператора - в папке data\out
pause
