@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Договоры 44-ФЗ членов СРО
echo с 01.07.2017 по вчерашний день, заказчики всех регионов.
echo Займёт несколько суток, можно запускать на ночь.
echo Окно можно закрыть: скачанные дни
echo при новом запуске не повторяются.
"%~dp0.venv\Scripts\python.exe" run.py history 44
if errorlevel 1 (
echo.
echo Загрузка остановилась с ошибкой. Запустите этот файл снова: скачанное не повторяется.
pause
exit /b 1
)
echo.
echo Готово. Сверка и Excel - в папке data\out.
echo Пришлите лог history_ДАТА.log из data\logs
pause
