@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Запускайте ПОСЛЕ того, как загрузили пакет в НОСТРОЙ.
echo Последний пакет nostroy_*.csv из data\out будет
echo отмечен как переданный: эти договоры не попадут
echo в следующий пакет, пока у них не изменится цена.
echo.
echo Чтобы продолжить, нажмите любую клавишу.
pause >nul
"%~dp0.venv\Scripts\python.exe" run.py mark-sent
pause
