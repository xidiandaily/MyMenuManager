@echo off
setlocal enabledelayedexpansion

:: ����Ƿ��ṩ���ļ�������
if "%~1"=="" (
    echo ���ṩ�ļ�����Ϊ������
    exit /b 1
)

:: ��ȡ������ļ���
set "filename=%~1"

:: �滻 \ Ϊ \\
set "escaped_filename="
for %%A in ("%filename:\=\\%") do set "escaped_filename=%%~A"

:: ��������Ƶ�������
echo|set /p="!escaped_filename!" | clip

:: ��ʾ���
echo �ļ����Ѹ��Ƶ�ճ���壺!escaped_filename!
exit /b 0
