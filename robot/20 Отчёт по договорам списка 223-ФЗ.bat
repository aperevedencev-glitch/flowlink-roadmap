@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Что случилось с каждым договором из списка.
echo Токен не нужен, ЕИС не запрашивается.
echo Только чтение скачанных архивов.
echo.
"%~dp0.venv\Scripts\python.exe" run.py audit223
echo.
echo Готово. Файл audit223_ДАТА.xlsx - в папке data\out
pause
