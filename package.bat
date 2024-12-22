@echo off
set RELEASE_DIR=bin\Release
set PACKAGE_DIR=package

:: 创建打包目录
if exist %PACKAGE_DIR% rd /s /q %PACKAGE_DIR%
mkdir %PACKAGE_DIR%
mkdir %PACKAGE_DIR%\example

:: 复制必要文件
copy %RELEASE_DIR%\MyMenuManagerUI.exe %PACKAGE_DIR%\
copy %RELEASE_DIR%\MyMenuManager.dll %PACKAGE_DIR%\
copy %RELEASE_DIR%\YamlDotNet.dll %PACKAGE_DIR%\
copy example\copy_path_and_escape.bat %PACKAGE_DIR%\example\
copy example\show_current_path.bat %PACKAGE_DIR%\example\
copy example\show_file_attributes.bat %PACKAGE_DIR%\example\
copy README.md %PACKAGE_DIR%\

:: 创建ZIP包
powershell Compress-Archive -Path %PACKAGE_DIR%\* -DestinationPath MyMenuManager.zip -Force

echo 打包完成：MyMenuManager.zip
pause 
