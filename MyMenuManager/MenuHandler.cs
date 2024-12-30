using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using YamlDotNet.Serialization;
using System.Collections.Generic;
using MyMenuManager.Models;
using Microsoft.Win32;
using System.Reflection;
using System.Linq;

namespace MyMenuManager
{
    [ComVisible(true)]
    [Guid("B54D41BD-C50B-4101-B9F7-58AD55A742E8")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyMenuManager.MenuHandler")]
    [ComDefaultInterface(typeof(IContextMenu))]
    public class MenuHandler : IContextMenu, IShellExtInit
    {
        private const string TYPE_FILE      = "file";
        private const string TYPE_DIRECTORY = "directory";
        private const string TYPE_BACKGROUD = "background";
        private const string DEFAULT_MENU_SHOW_DIR="查看配置和运行日志";

        private string _selectedPath; // 点击右键所在的位置
        private string  _type;         // 是否选中了文件or文件夹
        private bool   _is_init;      // 是否初始化成功
        private List<MenuConfig> _menuConfigs;
        private readonly string CONFIG_FILE;
        private Dictionary<uint, MenuConfig> _cmdMap; // 添加命令ID到MenuItem的映射
        private static readonly object _logLock = new object();
        private List<string> _defaultPaths;

        public MenuHandler()
        {
            string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string dllDirectory = Path.GetDirectoryName(dllPath);
            CONFIG_FILE = Path.Combine(dllDirectory, "config.yaml");
            _cmdMap = new Dictionary<uint, MenuConfig>();
            _defaultPaths = new List<string>();
            _defaultPaths.Add(dllDirectory);
            LoadConfig();
        }

        private void LoadConfig()
        {
            try
            {
                LogDebug("开始加载配置");
                
                // 获取默认菜单配置
                var defaultMenuConfigs = GetDefaultMenuConfigs();
                LogDebug($"已加载默认菜单配置: {defaultMenuConfigs.Count} 项");

                // 获取配置中的菜单
                var yamlMenuConfigs = GetMenuConfigs();

                // 更新菜单配置
                defaultMenuConfigs.AddRange(yamlMenuConfigs);
                _menuConfigs = defaultMenuConfigs;
                LogDebug($"配置加载完成，总计 {_menuConfigs.Count} 个菜单项");
            }
            catch (Exception ex)
            {
                LogDebug($"LoadConfig 异常: {ex.Message}");
                _menuConfigs = new List<MenuConfig>();
            }
        }

        private List<MenuConfig> LoadMenuConfigFromYaml(string yamlPath)
        {
            try
            {
                if (!File.Exists(yamlPath))
                {
                    LogDebug($"源配置文件不存在: {yamlPath}");
                    return new List<MenuConfig>();
                }

                string yamlDirectory = Path.GetDirectoryName(yamlPath);
                if (!_defaultPaths.Contains(yamlDirectory))
                {
                    _defaultPaths.Add(yamlDirectory);
                    LogDebug($"添加源配置文件目录到默认路径: {yamlDirectory}");
                }

                string yamlContent = File.ReadAllText(yamlPath);
                var deserializer = new DeserializerBuilder().Build();
                var configs = deserializer.Deserialize<List<MenuConfig>>(yamlContent);

                // 递归处理所有配置项的 source
                ProcessSourceConfigs(configs, new HashSet<string> { yamlPath });

                return configs;
            }
            catch (Exception ex)
            {
                LogDebug($"加载源配置文件失败 {yamlPath}: {ex.Message}");
                return new List<MenuConfig>();
            }
        }

        private void ProcessSourceConfigs(List<MenuConfig> configs, HashSet<string> loadedPaths)
        {
            if (configs == null) return;

            for (int i = 0; i < configs.Count; i++)
            {
                var config = configs[i];
                
                // 处理 source 字段
                if (!string.IsNullOrEmpty(config.Source))
                {
                    string sourcePath = GetAbsolutePath(config.Source);
                    if (!loadedPaths.Contains(sourcePath))
                    {
                        loadedPaths.Add(sourcePath);
                        var sourceConfigs = LoadMenuConfigFromYaml(sourcePath);
                        // 将加载的配置合并到当前位置
                        configs.InsertRange(i + 1, sourceConfigs);
                        i += sourceConfigs.Count; // 调整索引以跳过新插入的项
                    }
                }

                // 递归处理子菜单
                if (config.Submenu != null)
                {
                    ProcessSourceConfigs(config.Submenu, loadedPaths);
                }
            }
        }

        private string GetAbsolutePath(string path)
        {
            if (Path.IsPathRooted(path))
            {
                return path;
            }

            // 检查 PATH 中是否存在该命令
            string commandName = Path.GetFileName(path);
            string[] paths = Environment.GetEnvironmentVariable("PATH").Split(';');

            foreach (var cur_path in paths)
            {
                string fullPath = Path.Combine(cur_path, commandName);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            // 在所有默认路径中查找文件
            foreach (string basePath in _defaultPaths)
            {
                string fullPath = Path.Combine(basePath, path);
                if (File.Exists(fullPath))
                    return fullPath;
            }

            return path;
        }

        private StringBuilder GetFolderPath(IntPtr pidlFolder)
        {
            StringBuilder sbFolder = new StringBuilder(260);
            if (pidlFolder != IntPtr.Zero && Win32.SHGetPathFromIDList(pidlFolder, sbFolder))
            {
                return sbFolder;
            }
            return sbFolder;
        }

        public void Initialize(IntPtr pidlFolder, IntPtr dataObject, IntPtr hKeyProgID)
        {
            // 初始化
            _selectedPath = "";
            _type = "";
            _is_init = false;

            LogDebug("Initialize");

            try
            {
                // 如果有 dataObject，说明用户选中了文件/文件夹
                if (dataObject != IntPtr.Zero)
                {
                    var data = (System.Runtime.InteropServices.ComTypes.IDataObject)
                        Marshal.GetObjectForIUnknown(dataObject);

                    FORMATETC fe = new FORMATETC
                    {
                        cfFormat = (short)Win32.CF_HDROP,
                        ptd = IntPtr.Zero,
                        dwAspect = DVASPECT.DVASPECT_CONTENT,
                        lindex = -1,
                        tymed = TYMED.TYMED_HGLOBAL
                    };

                    STGMEDIUM stm;
                    data.GetData(ref fe, out stm);

                    var count = Win32.DragQueryFile(stm.unionmember, uint.MaxValue, null, 0);
                    if (count > 0)
                    {
                        var sb = new StringBuilder(260);
                        Win32.DragQueryFile(stm.unionmember, 0, sb, sb.Capacity);
                        _selectedPath = sb.ToString();
                        if (Directory.Exists(_selectedPath))
                        {
                            _type = TYPE_DIRECTORY;
                        }
                        else
                        {
                            _type = TYPE_FILE;
                        }
                        LogDebug($"选中的路径: {_selectedPath}, 类型: {_type}");
                        _is_init = true;
                    }

                    Marshal.ReleaseComObject(data);
                    Win32.ReleaseStgMedium(ref stm);
                }
                // 如果没有 dataObject 但有 pidlFolder，说明在文件夹空白处右键
                else if (pidlFolder != IntPtr.Zero)
                {
                    StringBuilder sbFolder = GetFolderPath(pidlFolder);
                    _selectedPath = sbFolder.ToString();
                    if (Directory.Exists(_selectedPath))
                    {
                        _type = TYPE_BACKGROUD;
                        _is_init = true;
                        LogDebug($"在文件夹空白处点击右键，文件夹路径:{_selectedPath}");
                    }
                }
                else
                {
                    LogDebug("not found any");
                }
            }
            catch (Exception ex)
            {
                LogDebug($"Initialize 失败: {ex}");
            }
        }

        public int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, 
            uint idCmdLast, uint uFlags)
        {
            try
            {
                _cmdMap.Clear(); // 清空之前的映射
                if (!_is_init){
                    LogDebug("pidlFolder 为空，没有必要生成菜单，直接返回");
                    return Win32.MAKE_HRESULT(Win32.SEVERITY_SUCCESS,0,1);
                }

                LogDebug($"开始创建菜单，参数：hMenu={hMenu}, indexMenu={indexMenu}, idCmdFirst={idCmdFirst}, idCmdLast={idCmdLast}, uFlags={uFlags}");
                LogDebug($"当前选中路径: {_selectedPath} 类型:{_type}");
                
                // 创建主菜单
                uint insertpos = 0;
                uint currentCmd = idCmdFirst;
                //bool mainMenuResult = Win32.InsertMenu(hMenu, indexMenu, Win32.MF_BYPOSITION | Win32.MF_STRING, currentCmd++, "my_menu");
                //LogDebug($"主菜单创建结果: {mainMenuResult}");

                // 创建子菜单
                IntPtr subMenu = Win32.CreatePopupMenu();
                LogDebug($"子菜单句柄: {subMenu}");
                
                // 判断当前是文件还是目录
                LogDebug($"目标类型: {_type}");
                
                bool is_add_item = false;
                foreach (var menuItem in _menuConfigs)
                {
                    LogDebug($"处理菜单项: {menuItem.Title}, 目标类型: {menuItem.Target}");
                    if (menuItem.Submenu != null && menuItem.Submenu.Count > 0)
                    {
                        if(AddMenuItem(subMenu, menuItem,idCmdFirst,ref insertpos, ref currentCmd))
                        {
                            is_add_item = true;
                        }
                    }
                    else if (menuItem.Source != null)
                    {
                        LogDebug($"加载Source:{menuItem.Source}");
                    }
                    else if (menuItem.Target != null && CheckTarget(menuItem.Target,_type))
                    {
                        if(IsValidCommand(menuItem.Cmd))
                        {
                            // 只添加匹配当前目标类型的菜单项
                            if(AddMenuItem(subMenu, menuItem,idCmdFirst,ref insertpos, ref currentCmd))
                            {
                                is_add_item = true;
                            }
                            LogDebug($"已添加菜单项: {menuItem.Title}");
                        }
                        else
                        {
                            LogDebug($"跳过匹配但是命令无效的菜单项: {menuItem.Title} Target:{menuItem.Target} Cmd:{menuItem.Cmd}");
                        }
                    }
                    else
                    {
                        LogDebug($"跳过不匹配的菜单项: {menuItem.Title} Target:{menuItem.Target}");
                    }
                }

                //// 将子菜单附加到主菜单
                //Win32.InsertMenu(hMenu, indexMenu++, Win32.MF_BYPOSITION | Win32.MF_STRING, currentCmd++, "test_invokeCmd");

                int addedItems = (int)(currentCmd-idCmdFirst);
                if (!is_add_item)
                {
                    Win32.DestroyMenu(subMenu);
                    return Win32.MAKE_HRESULT(Win32.SEVERITY_SUCCESS,0,1);
                }

                LogDebug($"共添加了 {addedItems} 个菜单项");

                // 将子菜单附加到主菜单
                bool attachResult = Win32.InsertMenu(hMenu, indexMenu, Win32.MF_BYPOSITION | Win32.MF_POPUP,
                    (uint)subMenu.ToInt32(), "my_menu");
                LogDebug($"子菜单附加结果: {attachResult}");
                LogDebug("菜单创建完成");
                return Win32.MAKE_HRESULT(Win32.SEVERITY_SUCCESS,0, (uint)addedItems+1);
            }
            catch (Exception ex)
            {
                LogDebug($"创建菜单失败: {ex}");
                LogDebug($"异常堆栈: {ex.StackTrace}");
                MessageBox.Show($"创建菜单失败: {ex.Message}");
                return Marshal.GetHRForException(ex);
            }
        }

        private bool AddMenuItem(IntPtr parentMenu, MenuConfig menuItem, uint idCmdFirst, ref uint insertpos, ref uint cmdId)
        {
            IntPtr subMenu = IntPtr.Zero;
            try
            {
                bool is_add_item = false;
                if (menuItem.Submenu != null && menuItem.Submenu.Count > 0)
                {
                    LogDebug($"创建子菜单: {menuItem.Title}");
                    subMenu = Win32.CreatePopupMenu();
                    uint sub_insertpos = 0;
                    foreach (var subItem in menuItem.Submenu)
                    {
                        if (subItem.Submenu != null && subItem.Submenu.Count > 0)
                        {
                            if(AddMenuItem(subMenu, subItem,idCmdFirst,ref sub_insertpos, ref cmdId))
                            {
                                is_add_item = true;
                            }
                        }
                        else if(CheckTarget(subItem.Target,_type))
                        {
                            if(IsValidCommand(subItem.Cmd))
                            {
                                if(AddMenuItem(subMenu, subItem,idCmdFirst,ref sub_insertpos, ref cmdId))
                                {
                                    is_add_item = true;
                                }
                            }
                            else
                            {
                                LogDebug($"跳过匹配但是命令无效的菜单项: {subItem.Title} Target:{subItem.Target} Cmd:{subItem.Cmd}");
                            }
                        }
                        else
                        {
                            LogDebug($"跳过不匹配的子菜单项: {subItem.Title}");
                        }
                    }
                    if (is_add_item == false)
                    {
                        Win32.DestroyMenu(subMenu);
                        return false;
                    }

                    bool result = Win32.InsertMenu(parentMenu, insertpos++, Win32.MF_BYPOSITION|Win32.MF_POPUP,
                        (uint)subMenu.ToInt32(), menuItem.Title);
                    LogDebug($"子菜单创建结果: {result}");
                    return true;
                }
                else
                {
                    bool result = Win32.InsertMenu(parentMenu, insertpos++, Win32.MF_BYPOSITION|Win32.MF_STRING,
                        cmdId, menuItem.Title);
                    if (result)
                    {
                        _cmdMap[cmdId-idCmdFirst] = menuItem; // 添加命令ID到MenuItem的映射
                        LogDebug($"添加命令映射: cmdId={cmdId}, menuItem.Title={menuItem.Title}");
                    }
                    LogDebug($"菜单项添加结果: {result}, cmdId={cmdId}");
                    cmdId++;
                    return true;
                }
            }
            catch (Exception ex)
            {
                if (subMenu != IntPtr.Zero)
                {
                    Win32.DestroyMenu(subMenu);
                }
                LogDebug($"添加菜单项失败: {menuItem.Title}, 错误: {ex}");
                throw;
            }
        }

        private string GetDefaultCommand(string cmd)
        {
            string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string dllDirectory = Path.GetDirectoryName(dllPath);

            string fullPath = Path.Combine(dllDirectory,cmd);
            if (File.Exists(fullPath))
            {
                LogDebug("found Default Command:{fullPath}");
                return fullPath;
            }

            return cmd;
        }

        private bool IsValidCommand(string cmd)
        {
            try
            {
                string default_path = GetAbsolutePath(cmd);
                if (File.Exists(default_path))
                {
                    return true;
                }
                LogDebug($"命令文件不存在: {cmd} default_path:{default_path}");
                return false;
            }
            catch (Exception ex)
            {
                LogDebug($"验证命令失败: {ex.Message}");
                return false;
            }
        }

        private bool CheckTarget(string Target, string type)
        {
            return Target.Split(',').Select(t => t.Trim()).Contains(type);
        }

        public int InvokeCommand(IntPtr pici)
        {
            LogDebug("InvokeCommand");
            try
            {
                CMINVOKECOMMANDINFO ici = (CMINVOKECOMMANDINFO)Marshal.PtrToStructure(
                    pici, typeof(CMINVOKECOMMANDINFO));
                
                uint cmdId = (uint)ici.lpVerb.ToInt32();
                LogDebug($"InvokeCommand cmdId: {cmdId}");

                if (_cmdMap.TryGetValue(cmdId, out MenuConfig menuItem))
                {
                    if (!string.IsNullOrEmpty(menuItem.Cmd) && IsValidCommand(menuItem.Cmd))
                    {
                        ExecuteCommand(cmdId,menuItem, _selectedPath);
                    }
                    else
                    {
                        LogDebug($"菜单项 {menuItem.Title} 没有配置命令");
                    }
                }
                else
                {
                    LogDebug($"未找到命令ID {cmdId} 对应的菜单项");
                }
                return 0;
            }
            catch (Exception ex)
            {
                LogDebug($"执行命令失败: {ex.Message}");
                return Marshal.GetHRForException(ex);
            }
        }

        public int GetCommandString(uint idCmd, uint uFlags, IntPtr reserved, 
            StringBuilder commandString, uint cchMax)
        {
            // 实现帮助文本显示（可选）
            return 0;
        }

        private MenuConfig GetMenuItemByIndex(List<MenuConfig> items, int index)
        {
            int currentIndex = 0;
            return FindMenuItem(items, index, ref currentIndex);
        }

        private MenuConfig FindMenuItem(List<MenuConfig> items, int targetIndex, ref int currentIndex)
        {
            foreach (var item in items)
            {
                if (currentIndex == targetIndex)
                    return item;
                
                currentIndex++;
                
                if (item.Submenu != null)
                {
                    var result = FindMenuItem(item.Submenu, targetIndex, ref currentIndex);
                    if (result != null)
                        return result;
                }
            }
            return null;
        }

        private void LogDebug(string message)
        {
            try
            {
                lock (_logLock)
                {
                    string logPath = Path.Combine(Path.GetDirectoryName(CONFIG_FILE), "debug.log");
                    File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
                }
                //Debug.WriteLine(message);
            }
            catch { }
        }

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                string clsid = t.GUID.ToString("B");

                // 注册文件右键菜单
                string fileKeyPath = @"*\shellex\ContextMenuHandlers\MyMenuManager";
                using (RegistryKey fileKey = Registry.ClassesRoot.CreateSubKey(fileKeyPath))
                {
                    fileKey.SetValue(null, clsid);
                }

                // 注册文件夹右键菜单
                string dirKeyPath = @"Directory\shellex\ContextMenuHandlers\MyMenuManager";
                using (RegistryKey dirKey = Registry.ClassesRoot.CreateSubKey(dirKeyPath))
                {
                    dirKey.SetValue(null, clsid);
                }

                // 注册文件夹背景右键菜单
                string bgKeyPath = @"Directory\Background\shellex\ContextMenuHandlers\MyMenuManager";
                using (RegistryKey bgKey = Registry.ClassesRoot.CreateSubKey(bgKeyPath))
                {
                    bgKey.SetValue(null, clsid);
                }

                Debug.WriteLine("Shell extension registered successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error registering shell extension: {ex.Message}");
                throw;  // 重新抛出异常，让注册过程失败
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                Debug.WriteLine($"Unregister CLSID: {t.GUID.ToString("B")}");

                // 删除文件右键菜单
                Registry.ClassesRoot.DeleteSubKeyTree(@"*\shellex\ContextMenuHandlers\MyMenuManager", false);

                // 删除文件夹右键菜单
                Registry.ClassesRoot.DeleteSubKeyTree(@"Directory\shellex\ContextMenuHandlers\MyMenuManager", false);

                // 删除文件夹背景右键菜单
                Registry.ClassesRoot.DeleteSubKeyTree(@"Directory\Background\shellex\ContextMenuHandlers\MyMenuManager", false);

                Debug.WriteLine("Shell extension unregistered successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error unregistering shell extension: {ex.Message}");
                // 注销过程中的错误可以忽略，因为可能是键已经不存在
            }
        }

        private List<MenuConfig> GetDefaultMenuConfigs()
        {
            string exePath = Assembly.GetExecutingAssembly().Location;
            string exeDir = Path.GetDirectoryName(exePath);
            string uiExePath = Path.Combine(exeDir, "MyMenuManagerUI.exe");
            string logPath = Path.Combine(exeDir, "log.txt");
            
            // 构造默认菜单配置
            var defaultConfig = new MenuConfig
            {
                Title = "默认",
                Submenu = new List<MenuConfig>
                {
                    new MenuConfig
                    {
                        Title  = "运行 MyMenuManagerUI",
                        Cmd    = uiExePath,
                        Target = "file,directory,background"  // 在所有类型下都显示
                    },
                    new MenuConfig
                    {
                        Title  = DEFAULT_MENU_SHOW_DIR,
                        Cmd    = "explorer.exe",
                        Target = "file,directory,background"
                    }
                }
            };

            return new List<MenuConfig> { defaultConfig };
        }

        private List<MenuConfig> GetMenuConfigs()
        {
            // 如果存在配置文件，则加载并合并配置
            if (File.Exists(CONFIG_FILE))
            {
                try
                {
                    return LoadMenuConfigFromYaml(CONFIG_FILE);
                }
                catch (Exception ex)
                {
                    LogDebug($"加载YAML {CONFIG_FILE}配置文件失败: {ex.Message}");
                    return new List<MenuConfig>();
                }
            }
            else
            {
                LogDebug($"配置文件不存在: {CONFIG_FILE}");

                string dllDirectory = Path.GetDirectoryName(CONFIG_FILE);
                string importyaml_path = Path.Combine(dllDirectory, "importyaml_config.yaml");

                string yamlString = "";

                yamlString  ="- Title: \"导入配置示例\"\r\n";
                yamlString +="  Submenu:  \r\n";
                yamlString +="    - Title: \"综合示例_复制路径\"\r\n";
                yamlString +="      Target: \"background,file,directory\"\r\n";
                yamlString +="      Cmd: \"example\\\\copy_path_and_escape.bat\"\r\n";
                File.WriteAllText(importyaml_path, yamlString);
                LogDebug($"已创建新的被导入配置文件:{importyaml_path}");

                yamlString  ="- Title: \"示例\"\r\n";
                yamlString +="  Submenu:  \r\n";
                yamlString +="    - Title: \"文件菜单示例_使用NotePad打开文件\"\r\n";
                yamlString +="      Target: \"file\"\r\n";
                yamlString +="      Cmd: \"notepad.exe\"\r\n";
                yamlString +="    - Title: \"文件夹菜单示例_显示属性\"\r\n";
                yamlString +="      Target: \"directory\"\r\n";
                yamlString +="      Cmd: \"example\\\\show_file_attributes.bat\"\r\n";
                yamlString +="    - Title: \"文件空白处菜单示例_显示当前路径\"\r\n";
                yamlString +="      Target: \"background\"\r\n";
                yamlString +="      Cmd: \"example\\\\show_current_path.bat\"\r\n";
                yamlString +="- Title: \"ImportYaml\"\r\n";
                yamlString +="  Source: \"importyaml_config.yaml\"\r\n";
                File.WriteAllText(CONFIG_FILE, yamlString);
                LogDebug($"已创建新的配置文件:{CONFIG_FILE}");

                return GetMenuConfigs();
            }
        }

        private void ExecuteCommand(uint cmdId,MenuConfig menuItem, string selectedPath)
        {
            try
            {
                if (menuItem.Title == DEFAULT_MENU_SHOW_DIR)
                {
                    string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    string dllDirectory = Path.GetDirectoryName(dllPath);
                    selectedPath = dllDirectory;
                }

                string cmd = GetAbsolutePath(menuItem.Cmd);
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = cmd,
                             Arguments = $"\"{selectedPath}\"",
                             UseShellExecute = false,
                             CreateNoWindow = false,
                             RedirectStandardOutput = false,
                             RedirectStandardError = false
                };
                LogDebug($"启动信息: Menu.title={menuItem.Title} cmdId={cmdId} cmd={cmd} FileName={startInfo.FileName}, Arguments={startInfo.Arguments}, UseShellExecute={startInfo.UseShellExecute}");
                Process.Start(startInfo);

            }
            catch (Exception ex)
            {
                LogDebug($"执行命令失败: {ex.Message}");
            }
        }
    }
} 
