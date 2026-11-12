@echo off
chcp 936 >nul
cd /d "%~dp0"
call "%~dp0_runtime.bat"

if not defined PYEXE (
    echo.
    echo   没有找到 Python，程序无法启动。
    echo   需要 Python 3.9 以上版本；只做命令行检查不需要额外装包。
    echo.
    pause
    goto :eof
)

set "CODE=%~1"
if "%CODE%"=="" set /p CODE=请输入工程编号（例如 NC23A005）:
if "%CODE%"=="" goto end
"%PYEXE%" "%~dp0check.py" %CODE%
echo.
pause
:end
