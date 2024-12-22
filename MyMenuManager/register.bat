@echo off
cd /d %~dp0
echo 正在注册 COM 组件...
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase "bin\Debug\MyMenuManager.dll"
echo 正在导入注册表...
REM reg import register.reg
REM pause 
