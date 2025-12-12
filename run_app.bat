@echo off
setlocal

rem Launch the Image -> Word Link Builder GUI (double-click friendly)
pushd "%~dp0"

set "PYTHON_EXE=python"
if exist ".venv\Scripts\python.exe" (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
)

"%PYTHON_EXE%" app.py
set "EXIT_CODE=%ERRORLEVEL%"

popd
if %EXIT_CODE% neq 0 (
    echo.
    echo App exited with error %EXIT_CODE%. See messages above.
    pause
)
exit /b %EXIT_CODE%
