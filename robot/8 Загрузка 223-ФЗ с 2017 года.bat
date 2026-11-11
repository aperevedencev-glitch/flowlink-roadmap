@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Договоры 223-ФЗ членов СРО, все регионы,
echo с 01.07.2017 по вчерашний день.
echo Идёт по годам от свежих к старым, после
echo каждого года - сверка и дашборды.
echo Займёт несколько суток. Окно можно закрыть:
echo скачанные дни при новом запуске пропускаются.
echo Компьютер на это время не уснёт.
echo.
"%~dp0.venv\Scripts\python.exe" run.py history223
echo.
echo Готово. Сверка и Excel - в папке data\out
pause
