@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Повторный поиск ненайденных договоров, окно 60 дней.
echo Нужен токен ЕИС. Займёт около часа.
echo Окно можно закрыть: скачанное не повторяется.
echo Закройте decisions.xlsx перед запуском.
echo.
"%~dp0.venv\Scripts\python.exe" run.py missing223
echo.
echo Готово. Сверка - в папке data\out
pause
