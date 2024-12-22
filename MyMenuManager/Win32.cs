using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace MyMenuManager
{
    internal static class Win32
    {
        public const uint MF_BYPOSITION = 0x400;
        public const uint MF_STRING = 0x0;
        public const uint MF_POPUP = 0x10;
        public const uint CF_HDROP = 15;

        public const int S_OK = 0;
        public const int S_FALSE = 1;
        public const uint SEVERITY_SUCCESS = 0;
        public const uint SEVERITY_ERROR = 1;
        public const uint FACILITY_NULL = 0;
        public const uint FACILITY_WIN32 = 7;
        
        public static int MAKE_HRESULT(uint severity, uint facility, uint code)
        {
            return (int)((severity << 31) | (facility << 16) | code);
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern bool InsertMenu(IntPtr hMenu, uint uPosition, uint uFlags,
            uint uIDNewItem, string lpNewItem);

        [DllImport("user32.dll")]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll")]
        public static extern bool DestroyMenu(IntPtr hMenu);

        [DllImport("shell32.dll")]
        public static extern uint DragQueryFile(IntPtr hDrop, uint iFile,
            [Out] StringBuilder lpszFile, int cch);

        [DllImport("ole32.dll")]
        public static extern void ReleaseStgMedium(ref System.Runtime.InteropServices.ComTypes.STGMEDIUM pmedium);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern bool SHGetPathFromIDList(IntPtr pidl, StringBuilder pszPath);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct CMINVOKECOMMANDINFO
    {
        public int cbSize;          // 结构体的大小（以字节为单位）
        public uint fMask;          // 标志，指示结构体的特性
        public IntPtr hwnd;         // 调用命令的窗口句柄
        public IntPtr lpVerb;       // 要执行的命令的名称（通常是字符串）
        [MarshalAs(UnmanagedType.LPStr)]
        public string lpParameters; // 命令参数（如果有的话）
        [MarshalAs(UnmanagedType.LPStr)]
        public string lpDirectory;  // 当前目录（通常是命令执行的目录）
        public int nShow;           // 窗口显示方式（例如，是否最小化、最大化等）
        public int dwHotKey;        // 快捷键（如果有的话）
        public IntPtr hIcon;        // 命令的图标句柄
    }
} 
