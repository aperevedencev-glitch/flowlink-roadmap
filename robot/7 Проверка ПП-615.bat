@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Проверяю, какие документы ПП-615 отдаёт ЕИС. Нужен токен ЕИС, займёт 5-15 минут.
"%~dp0.venv\Scripts\python.exe" run.py probe615
echo.
echo Готово. Пришлите файл probe615_*.log из папки data\logs
pause
