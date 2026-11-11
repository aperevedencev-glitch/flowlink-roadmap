@echo off
chcp 65001 >nul
cd /d "%~dp0"
call :setup || goto :end
"%~dp0.venv\Scripts\python.exe" run.py probe
:end
echo.
pause
exit /b

:setup
if exist "%~dp0.venv\Scripts\python.exe" goto :deps
set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY (
  echo Не найден Python. Установите его с https://www.python.org/downloads/ ^(галочка "Add python.exe to PATH"^) и запустите файл снова.
  exit /b 1
)
echo Готовлю окружение, это один раз и 1-2 минуты...
%PY% -m venv "%~dp0.venv" || exit /b 1
:deps
"%~dp0.venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt || exit /b 1
exit /b 0
