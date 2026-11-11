@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Точечная загрузка договоров 44-ФЗ по списку из папки data.
echo Перед запуском закройте окно файла 14.
echo Окно можно закрыть: найденное не повторяется.
echo.
"%~dp0.venv\Scripts\python.exe" run.py targeted44
if errorlevel 1 (
echo.
echo Загрузка остановилась с ошибкой. Запустите этот файл снова: скачанное не повторяется.
pause
exit /b 1
)
echo.
echo Готово. Отчёт targeted44_ДАТА.xlsx и сверка - в папке data\out.
echo Затем снова запустите файл 14.
pause
