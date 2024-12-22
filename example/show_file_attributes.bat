@echo off
:: 检查是否提供了文件名参数
if "%~1"=="" (
    echo 请输入文件名作为参数！
    echo 使用方式: %~nx0 文件名
    exit /b 1
)

:: 保存文件名
set "filename=%~1"

:: 检查文件是否存在
if not exist "%filename%" (
    echo 文件 "%filename%" 不存在！
    exit /b 1
)

:: 显示文件属性
echo 文件属性：
attrib "%filename%"

:: 显示文件详细信息
echo 文件详细信息：
for %%i in ("%filename%") do (
    echo 文件路径: %%~fi
    echo 文件名: %%~nxi
    echo 文件大小: %%~zi 字节
    echo 修改时间: %%~ti
)

pause