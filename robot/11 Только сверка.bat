@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Сверка уже найденных договоров по действующим правилам. Токен не нужен, 1-2 минуты.
"%~dp0.venv\Scripts\python.exe" run.py reconcile
echo.
echo Готово. Excel для оператора - в папке data\out
pause
