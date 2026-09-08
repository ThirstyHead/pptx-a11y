@echo off
rem Build standalone Windows executable using PyInstaller
echo ==> Building pptx-a11y Windows executable...
call .venv\Scripts\activate.bat
pyinstaller --clean -y packaging\pptx-a11y.spec
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller build failed.
    exit /b %ERRORLEVEL%
)
echo ==> Windows executable built at: dist\pptx-a11y\pptx-a11y.exe
