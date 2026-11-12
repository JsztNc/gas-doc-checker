@echo off
rem ---------------------------------------------------------------
rem 由 检查资料.bat / 命令行检查.bat 调用，探测本机可用的 Python
rem 成功后设置：PYEXE=python.exe  PYEXE 为空表示没找到
rem 注意：PATH 里 %LOCALAPPDATA%\Microsoft\WindowsApps\python.exe 是应用商店
rem 的占位程序，双击只会弹商店，必须排除掉
rem ---------------------------------------------------------------
set "PYEXE="
set "PYWEXE="

rem 1) PATH 里的 python.exe（排除 WindowsApps 占位）
for %%I in (python.exe) do (
    if not "%%~$PATH:I"=="" (
        echo %%~$PATH:I| findstr /i /c:"WindowsApps" >nul || set "PYEXE=%%~$PATH:I"
    )
)

rem 2) WorkBuddy 自带的解释器
if not defined PYEXE (
    for /d %%D in ("%USERPROFILE%\.workbuddy\binaries\python\versions\*") do (
        if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)

rem 3) 用户目录下的常规安装
if not defined PYEXE (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python*") do (
        if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)

rem 4) Program Files 与根目录安装
if not defined PYEXE (
    for /d %%D in ("%ProgramFiles%\Python*") do (
        if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)
if not defined PYEXE (
    for /d %%D in ("C:/Python*") do (
        if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)

rem 由 python.exe 推出同目录的 pythonw.exe（无黑窗启动）
if defined PYEXE (
    for %%P in ("%PYEXE%") do (
        if exist "%%~dpPpythonw.exe" set "PYWEXE=%%~dpPpythonw.exe"
    )
)
