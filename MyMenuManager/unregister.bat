@echo off
cd /d %~dp0
echo ���ڽ�����ؽ���...
taskkill /F /IM explorer.exe
taskkill /F /IM "Q-Dir_x64.exe"
timeout /t 2
echo ����ע�� COM ���...
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /u "bin\Debug\MyMenuManager.dll"
echo ����ɾ��ע�����...
REM reg delete "HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\MyMenuManager" /f
REM reg delete "HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers\MyMenuManager" /f
REM reg delete "HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers\MyMenuManager" /f
REM reg delete "HKEY_CLASSES_ROOT\CLSID\{B54D41BD-C50B-4101-B9F7-58AD55A742E8}" /f
REM reg delete "HKEY_CLASSES_ROOT\MyMenuManager.MenuHandler" /f
REM reg delete "HKEY_CLASSES_ROOT\Record\{B54D41BD-C50B-4101-B9F7-58AD55A742E8}\1.0.0.0" /f
echo ����������Դ������...
start explorer.exe
start D:\Q-Dir_Portable_x64\Q-Dir\Q-Dir_x64.exe
REM pause 
