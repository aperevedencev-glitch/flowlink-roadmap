@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Предмет договоров 223-ФЗ из извещений ЕИС.
echo Нужен токен ЕИС. Займёт несколько минут.
echo Договоры со строительным предметом уйдут к передаче,
echo текущий ремонт и содержание будут исключены.
echo Закройте decisions.xlsx перед запуском.
echo.
"%~dp0.venv\Scripts\python.exe" run.py subjects223
echo.
echo Готово. Сверка - в папке data\out
pause
