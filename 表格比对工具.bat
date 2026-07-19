@echo off
chcp 65001 >nul
title 表格比对工具 v2.0.0
powershell -ExecutionPolicy Bypass -File "%~dp0compare.ps1"
