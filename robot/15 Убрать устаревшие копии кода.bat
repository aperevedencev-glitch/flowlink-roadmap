@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Переношу устаревшие копии кода в папку
echo "_устаревшие_копии_кода". Ничего не удаляется.
echo Робот работает из папки eis2nostroy.
echo Не запускайте, пока идёт загрузка или пересборка.
echo.
echo Чтобы продолжить, нажмите любую клавишу.
pause >nul
set OLD=_устаревшие_копии_кода
if not exist "%OLD%" mkdir "%OLD%"
if not exist "%OLD%\eis2nostroy" mkdir "%OLD%\eis2nostroy"
for %%F in (__init__.py __main__.py alerts.py cli.py config.py eis_client.py export.py members.py models.py nostroy.py parser.py protocols.py reconcile.py sources.py tabular.py) do (
  if exist "%%F" move /Y "%%F" "%OLD%\" >nul
)
if exist "eis2nostroy\run.py" move /Y "eis2nostroy\run.py" "%OLD%\eis2nostroy\" >nul
if exist "eis2nostroy\config.toml" move /Y "eis2nostroy\config.toml" "%OLD%\eis2nostroy\" >nul
if exist "eis2nostroy\*.bat" move /Y "eis2nostroy\*.bat" "%OLD%\eis2nostroy\" >nul
if exist "eis2nostroy\eis2nostroy" move /Y "eis2nostroy\eis2nostroy" "%OLD%\eis2nostroy\eis2nostroy_old" >nul
echo.
echo Готово. Проверка, что робот запускается:
"%~dp0.venv\Scripts\python.exe" run.py tests
pause
