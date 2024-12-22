@echo off
setlocal enabledelayedexpansion

:: 检查是否提供了文件名参数
if "%~1"=="" (
    echo 请提供文件名作为参数。
    exit /b 1
)

:: 获取传入的文件名
set "filename=%~1"

:: 替换 \ 为 \\
set "escaped_filename="
for %%A in ("%filename:\=\\%") do set "escaped_filename=%%~A"

:: 将结果复制到剪贴板
echo|set /p="!escaped_filename!" | clip

:: 显示结果
echo 文件名已复制到粘贴板：!escaped_filename!
exit /b 0
