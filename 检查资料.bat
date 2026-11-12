@echo off
chcp 936 >nul
cd /d "%~dp0"
call "%~dp0_runtime.bat"

if defined PYWEXE (
    start "" "%PYWEXE%" "%~dp0ui.py"
    goto :eof
)
if defined PYEXE (
    start "" "%PYEXE%" "%~dp0ui.py"
    goto :eof
)

echo.
echo   没有找到 Python，程序无法启动。
echo.
echo   本工具需要 Python 3.9 以上版本，并且要装 pythonnet：
echo       pip install pythonnet pywin32
echo.
echo   装好后重新双击本文件即可。
echo.
pause
