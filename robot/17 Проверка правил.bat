@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Автотесты робота и проверка правил
echo на решениях оператора (data\decisions.xlsx).
"%~dp0.venv\Scripts\python.exe" run.py tests
"%~dp0.venv\Scripts\python.exe" run.py etalon
echo.
echo Результат эталона - файл etalon_ДАТА.xlsx в data\out
pause
