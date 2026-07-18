@echo off
chcp 65001 >nul
title 表格比对工具 v2.0.0
echo.
echo   +--------------------------------------+
echo   ^|    表格比对工具 v2.0.0             ^|
echo   ^|    Excel 逐列比对 - 差异标红      ^|
echo   +--------------------------------------+
echo.
echo   正在启动，请在弹出的窗口中操作...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0compare.ps1"
