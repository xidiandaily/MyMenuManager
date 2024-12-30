# MyMenuManager

``my_menu``，Windows右键菜单自定义工具，其中顶层目录``my_menu``是默认添加的，在``my_menu``下面的二层、三层菜单都可以通过``yaml``配置文件进行自定义配置，并且对菜单对应命令进行响应。

``yaml``配置文件可以配置该菜单是在 ``file``(文件)、``directory``（目录）、``background``（文件夹背景菜单） 进行显示。

通过这个自定义工具，可以在windows上构造属于自己的效率工具。

## 功能特性

- 提供简单的图形界面进行安装/卸载
- 支持自定义右键菜单项
- 支持 Windows 资源管理器集成
- 自动记录安装/卸载日志

## 使用方法
> 这里说明一下工具的使用方法

整个``my_menu``的菜单结构都是通过配置进行构造出来的，其菜单、菜单对应的命令都是通过``yaml``格式进行配置。

一个菜单项的配置方法如下:
```yaml

- title: "文件菜单示例_使用NotePad打开文件"
  target: "file"
  cmd: "notepad.exe"

```

title
: 这个是显示在右键菜单上的名称
target
: 这个说明这个菜单项是在那种类型对象上点击右键时显示，取值范围 ``file``、``directory``、``background``
: - ``file`` 在文件上点击右键时会显示，其他类型对不会显示
: - ``directory`` 在文件夹上点击右键会显示
: - ``background`` 在文件夹背景处点击右键可以显示
: - 可以采用 ',' 分割符进行复选，例如 "file,directory" 在文件和文件夹都会显示这个菜单项
cmd
: 这个是选中菜单项后，会执行的命令，并且会将当前的文件对象完整路径传给这个命令,命令支持的对象是 ``.exe``、``.bat``、``.cmd``
: 假如``cmd``配置的是 ``notepad.exe``， 那么就会使用 ``notepad.exe``将文件打开
: 这里的 ``cmd``可以自己定义，可以是``notepad.exe``、``gvime.exe``、甚至是python脚本
subcmd
: 二级菜单，配置二级菜单的时候，本菜单项可以不配置``target``参数，以``subcmd``的配置为准
Source
:1，修改配置，原来每一个``MenuConfig``仅支持``Title``,``Target``,``Cmd``,``Submen``，现在新增加一个可选的新字段 ``Source``
:2, ``Source`` 字段对象是一个配置在另一个位置的 ``yaml``格式的 ``MenuConfig``， 当``LoadConfig`` 函数读到这个字段时，要自动加载这个``yaml``，并将其内容``merge``进来；
:3，增加多一个 ``default_paths`` 的变量，存放一些默认的路径，目前放到这个 ``default_paths`` 中的路径有:
:   3.1 ``dllPath``所在的的 ``dllDirectory``
:   3.2 每一个``Source``字段对应配置 ``yaml`` 的 ``yamlDirectory``
:4, ``default_paths`` 中的优先级，按照 ``dllPath``,``yamlDirectory`` 来，查找可执行命令时，优先查找``PATH``,再查找 ``default_paths``

### 配置案例
> 示例配置说明
示例配置是在将 ``my_menu``导入到右键菜单之后，首次呼出右键菜单时，会在``MyMenuManager.dll`` 所在路径下查找默认配置 ``config.yaml``,如果没有找到，就会生成如下面所示的示例 ``config.yaml``。

```yaml
//config.yaml
- Title: "示例"
  Submenu:  
    - Title: "文件菜单示例_使用NotePad打开文件"
      Target: "file"
      Cmd: "notepad.exe"
    - Title: "文件夹菜单示例_显示属性"
      Target: "directory"
      Cmd: "example\\show_file_attributes.bat"
    - Title: "文件空白处菜单示例_显示当前路径"
      Target: "background"
      Cmd: "example\\show_current_path.bat"
    - Title: "综合示例_复制路径"
      Target: "background,file,directory"
      Cmd: "example\\copy_path_and_escape.bat"
- Title: "ImportYaml"
  Source: "importyaml_config.yaml"

//file
- Title: "导入配置示例"
  Submenu:  
    - Title: "综合示例_复制路径"
      Target: "background,file,directory"
      Cmd: "example\\copy_path_and_escape.bat"

```
其呼出的菜单如下所示:
1. 在文件上呼出右键菜单
```txt
my_menu->"示例"->文件菜单示例_使用NotePad打开文件
               ->综合示例_复制路径
            
```
2. 在文件夹上呼出右键菜单
```txt
my_menu->"示例"->文件菜单示例_使用NotePad打开文件
               ->综合示例_复制路径
            
```
3. 在文件夹背景上呼出右键菜单
```txt
my_menu->"示例"->文件菜单示例_使用NotePad打开文件
               ->综合示例_复制路径
            
```

![my_menu_running](https://github.com/user-attachments/assets/a23bef96-f384-4cc1-b32b-17c8d05caa1a)



## 项目结构

```txt
MyMenuManager/
├── MyMenuManager/           # 右键菜单处理程序
│   └── MenuHandler.cs      # COM 组件实现
├── MyMenuManagerUI/        # 图形界面程序
│   ├── MainForm.cs        # 主窗体
│   ├── Program.cs         # 程序入口
│   └── MainForm.resx      # 窗体资源
└── package.bat            # 打包脚本
```

## 系统要求

- Windows 7 或更高版本
- .NET Framework 4.0 或更高版本
- 需要管理员权限进行安装/卸载

## 安装和卸载说明

1. 下载并解压 `MyMenuManager.zip`
2. 以管理员身份运行 `MyMenuManagerUI.exe`
3. 点击“导入菜单”按钮将 ``my_menu``导入菜单
4. 安装完成后，可以在文件右键菜单中看到新增的功能
5. 点击“从右键菜单删除”按钮将 ``my_menu``从菜单删除
5. 卸载后，将看不到 ``my_menu``

![bandicam](https://github.com/user-attachments/assets/b896ec8d-0a6a-4ea3-9f74-1733e0316f12)


## 开发说明

### 技术栈
- C# WinForms
- COM 组件开发
- Windows Shell 扩展

### 关键文件说明
- `MenuHandler.cs`: 实现 Shell 扩展接口，处理右键菜单逻辑
- `MainForm.cs`: 提供图形界面，处理安装/卸载逻辑
- `YamlDotNet`: 用于配置文件解析

### 编译说明
1. 使用 Visual Studio 打开解决方案
2. 编译 MyMenuManager 项目生成 DLL
3. 编译 MyMenuManagerUI 项目生成主程序
4. 运行 `package.bat` 生成发布包

## 日志说明

UI程序运行时会在程序目录下生成 `ui_log.txt` 日志文件，记录安装/卸载过程的详细信息，便于问题排查。

dll程序运行时也会在程序目录下生成 ``log.txt``日志文件


## 注意事项

1. 安装和卸载都需要管理员权限
2. 如果安装失败，请检查：
   - 是否以管理员身份运行
   - .NET Framework 版本是否正确
   - 相关 DLL 文件是否完整
3. 程序会自动检测是否已安装，避免重复安装
4. 卸载时会自动刷新资源管理器

## 依赖文件

发布包中包含以下必要文件：
- `MyMenuManagerUI.exe`: 主程序
- `MyMenuManager.dll`: 右键菜单处理程序
- `YamlDotNet.dll`: YAML 解析库

## 问题反馈

如果遇到问题，请检查以下内容：
1. 点击 ``my_menu``->``默认``->``查看配置和运行日志``,查看 ``debug.log``,  `ui_log.txt` 日志文件.
2. 检查 ``config.yaml``格式是否正常
3. 检查 ``config.yaml``中的配置项执行文件是否存在，如果不存在，请添加完整路径
4. 检查所需文件是否完整

## 开发计划

- [x] 完成基本支持框架
- [x] 通过UI界面进行右键菜单的“导入”和“卸载”
- [x] 添加“默认”菜单
- [x] 添加“示例”命令
- [ ] 添加配置文件功能
- [ ] 支持更多自定义选项
- [ ] 添加多语言支持
