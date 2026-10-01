@echo off
echo 正在启动工作日报助手 (Fluent Design版)...
echo.

REM 切换到项目目录
cd /d "%~dp0"

REM 使用 daily-AI 环境的 Python 直接运行
"D:\ProgramSoftware\anaconda3\envs\daily-AI\python.exe" "UI\main_fluent.py"

REM 如果程序退出，暂停查看输出
pause
