@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Какую роль член СРО играл в договорах 44-ФЗ, выпавших при пересборке.
echo Токен не нужен, ЕИС не запрашивается.
echo Только чтение скачанных архивов.
echo.
"%~dp0.venv\Scripts\python.exe" run.py roles44
echo.
echo Готово. Файл roles44_ДАТА.xlsx - в папке data\out
pause
