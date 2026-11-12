@echo off
chcp 65001 >nul
title 解除网络标记
cd /d "%~dp0"

echo 正在解除本文件夹所有文件的"来自网络"标记…
echo （程序从网盘/邮件下载后，Windows 会锁住里面的 DLL 导致闪退）
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-ChildItem -LiteralPath '%~dp0' -Recurse -File | ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction Stop } catch {} }; Write-Host '完成。'"

echo.
echo 处理完成，现在可以双击「线上资料检查.exe」打开了。
echo 如果还是闪退，右键本文件夹 → 属性 → 勾选「解除锁定」→ 确定。
echo.
pause
