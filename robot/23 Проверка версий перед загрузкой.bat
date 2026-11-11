@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Проверка перед загрузкой в НОСТРОЙ: не устарела ли запись о договоре.
echo Токен не нужен, ЕИС не запрашивается, рабочий файл не меняется.
echo Просмотр всех скачанных архивов занимает от нескольких минут до часа.
echo.
"%~dp0.venv\Scripts\python.exe" run.py versions
echo.
echo Готово. Файл versions_ДАТА.xlsx - в папке data\out
pause
