using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

namespace MyMenuManagerUI
{
    public partial class MainForm : Form
    {
        private readonly string _dllPath;
        private readonly string _regasmPath;
        private readonly string _logPath;

        public MainForm()
        {
            InitializeComponent();
            
            try
            {
                // 获取 MyMenuManager.dll 的路径
                string exePath = Assembly.GetExecutingAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);
                _dllPath = Path.Combine(exeDir, "MyMenuManager.dll");
                _logPath = Path.Combine(exeDir, "ui_log.txt");

                // 获取 RegAsm 路径 - 修正路径
                _regasmPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"Microsoft.NET\Framework64\v4.0.30319\regasm.exe");

                // 检查文件是否存在
                CheckFiles();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化失败：{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void CheckFiles()
        {
            if (!File.Exists(_dllPath))
            {
                throw new FileNotFoundException(
                    $"找不到 MyMenuManager.dll 文件！\n期望路径：{_dllPath}");
            }

            if (!File.Exists(_regasmPath))
            {
                throw new FileNotFoundException(
                    $"找不到 RegAsm.exe！\n期望路径：{_regasmPath}");
            }
        }

        private bool IsDllRegistered()
        {
            try
            {
                // 检查 CLSID 注册表项
                string dllName = Path.GetFileNameWithoutExtension(_dllPath);
                string subKeyPath = $"*\\shellex\\ContextMenuHandlers\\MyMenuManager";
                LogDebug($"dllName{dllName} subKeyPath:{subKeyPath}");
                using (RegistryKey clsidKey = Registry.ClassesRoot.OpenSubKey($@"{subKeyPath}", false))
                {
                    LogDebug($"clsidKey:{clsidKey}");
                    return clsidKey != null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"检查 DLL 注册状态时出错: {ex.Message}");
                return false;
            }
        }

        private void LogDebug(string message)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                File.AppendAllText(_logPath, logEntry);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"写入日志失败: {ex.Message}");
            }
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            try
            {
                LogDebug("开始安装过程");
                LogDebug($"DLL路径: {_dllPath}");
                LogDebug($"RegAsm路径: {_regasmPath}");

                // 再次检查文件
                CheckFiles();
                LogDebug("文件检查通过");

                // 检查是否已安装
                if (IsDllRegistered())
                {
                    LogDebug("DLL已经注册，退出安装");
                    MessageBox.Show("右键菜单已经安装，无需重复安装！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Exit();
                    return;
                }

                // 注册 COM 组件
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = _regasmPath,
                    Arguments = $"/codebase \"{_dllPath}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                LogDebug("开始执行RegAsm注册");
                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    LogDebug($"RegAsm执行完成，退出代码: {process.ExitCode}");
                    if (process.ExitCode == 0)
                    {
                        LogDebug("安装成功");
                        MessageBox.Show("右键菜单安装成功！", "成功", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Application.Exit();
                    }
                    else
                    {
                        LogDebug($"安装失败，退出代码: {process.ExitCode}");
                        MessageBox.Show($"右键菜单安装失败！\n退出代码：{process.ExitCode}", 
                            "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }

                LogDebug("开始刷新资源管理器");
                RefreshExplorer();
                LogDebug("安装过程完成");
            }
            catch (Exception ex)
            {
                LogDebug($"安装过程出错: {ex.Message}");
                LogDebug($"异常详细信息: {ex}");
                MessageBox.Show($"安装过程出错：\n{ex.Message}\n\n详细信息：{ex}", 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void btnUninstall_Click(object sender, EventArgs e)
        {
            try
            {
                LogDebug("开始卸载过程");
                LogDebug($"DLL路径: {_dllPath}");
                LogDebug($"RegAsm路径: {_regasmPath}");

                // 再次检查文件
                CheckFiles();
                LogDebug("文件检查通过");

                // 检查是否已安装
                if (!IsDllRegistered())
                {
                    LogDebug("DLL未注册，退出卸载");
                    MessageBox.Show("右键菜单尚未安装，无需卸载！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Exit();
                    return;
                }

                // 注销 COM 组件
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = _regasmPath,
                    Arguments = $"/u \"{_dllPath}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                LogDebug("开始执行RegAsm卸载");
                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    LogDebug($"RegAsm执行完成，退出代码: {process.ExitCode}");
                    if (process.ExitCode == 0)
                    {
                        LogDebug("卸载成功");
                        MessageBox.Show("右键菜单卸载成功！", "成功", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Application.Exit();
                    }
                    else
                    {
                        LogDebug($"卸载失败，退出代码: {process.ExitCode}");
                        MessageBox.Show($"右键菜单卸载失败！\n退出代码：{process.ExitCode}", 
                            "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }

                LogDebug("开始刷新资源管理器");
                RefreshExplorer();
                LogDebug("卸载过程完成");
            }
            catch (Exception ex)
            {
                LogDebug($"卸载过程出错: {ex.Message}");
                LogDebug($"异常详细信息: {ex}");
                MessageBox.Show($"卸载过程出错：\n{ex.Message}\n\n详细信息：{ex}", 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void RefreshExplorer()
        {
            try
            {
                // 通知系统更新
                SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"刷新 Explorer 失败: {ex.Message}");
            }
        }

        // 添加 P/Invoke 声明
        [System.Runtime.InteropServices.DllImport("shell32.dll")]
        private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
    }
} 
