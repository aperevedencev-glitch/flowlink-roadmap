@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\python.exe" (echo Сначала запустите "1 Проверка токена.bat" & pause & exit /b 1)
echo Догружаю дни 2024 года, на которых ЕИС отвечал ошибкой: 22-24.02.2024 и 13.10.2024 (223-ФЗ).
echo Займёт около часа. Потом сверка пересоберётся по новым правилам.
"%~dp0.venv\Scripts\python.exe" run.py ranges 223 2024-02-22:2024-02-24 2024-10-13:2024-10-13
echo.
echo Готово. Пришлите файл лога из папки data\logs (ranges_ДАТА.log)
pause
