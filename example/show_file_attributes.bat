@echo off
:: ����Ƿ��ṩ���ļ�������
if "%~1"=="" (
    echo �������ļ�����Ϊ������
    echo ʹ�÷�ʽ: %~nx0 �ļ���
    exit /b 1
)

:: �����ļ���
set "filename=%~1"

:: ����ļ��Ƿ����
if not exist "%filename%" (
    echo �ļ� "%filename%" �����ڣ�
    exit /b 1
)

:: ��ʾ�ļ�����
echo �ļ����ԣ�
attrib "%filename%"

:: ��ʾ�ļ���ϸ��Ϣ
echo �ļ���ϸ��Ϣ��
for %%i in ("%filename%") do (
    echo �ļ�·��: %%~fi
    echo �ļ���: %%~nxi
    echo �ļ���С: %%~zi �ֽ�
    echo �޸�ʱ��: %%~ti
)

pause