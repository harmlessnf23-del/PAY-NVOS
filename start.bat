@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "app.py" (
  echo [ОШИБКА] Рядом нет app.py. Сначала распакуйте архив целиком в папку
  echo          ^(ПКМ по zip - "Извлечь все"^), потом запускайте start.bat оттуда.
  pause
  exit /b 1
)

rem --- ищем Python ---
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY (
  python -c "import sys" >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo [ОШИБКА] Python не найден. Установите Python 3.10+ с python.org
  echo          и отметьте галочку "Add python.exe to PATH".
  pause
  exit /b 1
)

rem --- ставим зависимости, если их нет ---
%PY% -c "import flask, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Устанавливаю зависимости ^(flask, openpyxl^)...
  %PY% -m pip install --user flask openpyxl
  if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости. Проверьте интернет/прокси.
    pause
    exit /b 1
  )
)

rem --- браузер откроется, когда сервер реально поднимется ---
start "" /min powershell -NoProfile -WindowStyle Hidden -Command ^
 "for($i=0;$i -lt 60;$i++){try{Invoke-WebRequest http://127.0.0.1:8021/healthz -UseBasicParsing -TimeoutSec 1 ^| Out-Null; Start-Process 'http://127.0.0.1:8021'; break}catch{Start-Sleep -Milliseconds 500}}"

echo Сервер сверки: http://127.0.0.1:8021  ^(не закрывайте это окно^)
%PY% app.py

echo.
echo Сервер остановлен. Если выше ошибка - пришлите её скриншот.
pause
