@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Обновляю HTML-дашборды руководителя и сотрудника из последней сверки. Токен не нужен.
"%~dp0.venv\Scripts\python.exe" run.py dashboards
if /i "%~1"=="/auto" exit /b
echo.
echo Готово. Файлы в папке data\out:
echo   Flowlink — дашборд руководителя.html
echo   Flowlink — дашборд сотрудника.html
start "" "%~dp0data\out\Flowlink — дашборд руководителя.html"
pause
