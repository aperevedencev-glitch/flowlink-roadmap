@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Точечная загрузка 223-ФЗ по списку договоров
echo из файла «Договоры 223-ФЗ ... .xlsx» в папке data.
echo Робот запрашивает только регион заказчика и
echo 10 дней после даты заключения. Займёт 2-3 часа.
echo Окно можно закрыть: скачанное не повторяется.
echo Сначала закройте окно «8 Загрузка 223-ФЗ».
echo.
"%~dp0.venv\Scripts\python.exe" run.py targeted223
echo.
echo Готово. Сверка - в data\out, ненайденные -
echo в файле targeted223_не_найдены_ДАТА.xlsx
pause
