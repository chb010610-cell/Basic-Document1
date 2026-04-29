@echo off
chcp 65001 >nul
title ISO质量体系文档生成器 v2
echo ============================================
echo    ISO质量体系文档批量生成器 v2
echo    模板目录: Desktop\website后台 (12个文件)
echo    正在启动服务...
echo ============================================
echo.

cd /d "C:\Users\Administrator\WorkBuddy\20260429103845"

"C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe" server.py

echo.
echo 服务已停止运行。
pause
