@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Договоры ПП-615 (капремонт МКД) членов СРО
echo с 01.07.2017 по вчерашний день, заказчики всех регионов.
echo Займёт 2-3 суток, можно запускать на ночь.
echo Окно можно закрыть: скачанные дни
echo при новом запуске не повторяются.
"%~dp0.venv\Scripts\python.exe" run.py ranges 615 2017-07-01:2026-11-02
echo.
echo Готово. Сверка и Excel - в папке data\out.
echo Пришлите лог ranges_ДАТА.log из data\logs
pause
