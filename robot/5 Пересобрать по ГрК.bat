@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Пересборка всех скачанных архивов по текущим правилам.
echo Идёт несколько часов, токен не нужен.
echo Окно можно закрыть: рабочий список договоров
echo заменяется только в самом конце, а повторный
echo запуск продолжит с места остановки.
echo Компьютер на это время не уснёт.
echo.
"%~dp0.venv\Scripts\python.exe" run.py rebuild
echo.
echo Готово. Excel для оператора - в папке data\out
pause
