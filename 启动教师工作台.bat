@echo off
chcp 65001 >nul
title 成绩分析工具
cd /d "%~dp0"
echo.
echo   成绩分析工具 —— 正在启动
echo.

set PY=
where python >nul 2>nul && set PY=python
if not defined PY (
  where py >nul 2>nul && set PY=py
)
if not defined PY (
  echo   没有找到 Python。
  echo.
  echo   这个工具需要 Python 才能运行，装一次以后就不用再装：
  echo     https://www.python.org/downloads/
  echo     安装时记得勾选 Add Python to PATH
  echo.
  pause
  exit /b 1
)

echo   用 %PY% 启动，浏览器会自动打开。
echo   关闭这个黑窗口就等于关闭工具。
echo.
%PY% teacher_app.py
if errorlevel 1 (
  echo.
  echo   启动失败。把上面红色的报错信息发给技术老师。
  pause
)