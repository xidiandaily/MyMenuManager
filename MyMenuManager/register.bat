@echo off
cd /d %~dp0
echo ����ע�� COM ���...
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase "bin\Debug\MyMenuManager.dll"
echo ���ڵ���ע���...
REM reg import register.reg
REM pause 
