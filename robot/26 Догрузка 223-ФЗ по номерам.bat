@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Догрузка договоров 223-ФЗ по реестровым номерам.
echo Список: data\Догрузка 223-ФЗ по номерам.xlsx
echo Окно можно закрыть: найденное не повторяется.
echo.
"%~dp0.venv\Scripts\python.exe" run.py targeted223reg
if errorlevel 1 (
echo.
echo Загрузка остановилась с ошибкой. Запустите этот файл снова: скачанное не повторяется.
pause
exit /b 1
)
echo.
echo Готово. Отчёт targeted223_по_номерам_ДАТА.xlsx и сверка - в папке data\out.
pause
