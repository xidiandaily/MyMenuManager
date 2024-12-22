using System;
using System.Runtime.InteropServices;
using System.Text;

namespace MyMenuManager
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214E4-0000-0000-C000-000000000046")]
    public interface IContextMenu
    {
        [PreserveSig]
        int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);

        int InvokeCommand(IntPtr pici);

        [PreserveSig]
        int GetCommandString(uint idCmd, uint uFlags, IntPtr reserved, StringBuilder commandString, uint cchMax);
    }

    //这段代码定义了一个 COM 接口 IShellExtInit，它主要用于 Windows Shell 扩展的初始化过程。
    //开发者通过实现这个接口，能够创建右键菜单扩展或其他自定义的 Shell 功能。
    //比如，当用户右键单击文件夹时，Shell 扩展可能调用该接口的 Initialize 方法，将文件夹的信息传递给扩展程序，供其进行初始化。
    //Initialize 方法的调用场景主要包括以下几种：
    //   (1) 右键菜单扩展（Context Menu Handlers）
    //   当用户右键单击某个对象（文件或文件夹）时，Windows Shell 会加载右键菜单扩展。
    //   在加载扩展时，调用 Initialize 方法，将用户所选对象的信息（如文件路径、对象类型等）传递给扩展程序。
    //   扩展程序根据这些信息动态生成右键菜单项。
    //   (2) 属性表扩展（Property Sheet Handlers）
    //   当用户右键单击对象并选择“属性”时，Initialize 方法会被调用，提供对象的详细信息，以便扩展程序创建自定义属性页。
    //   (3) 自定义图标覆盖（Icon Overlay Handlers）
    //   在为文件或文件夹设置自定义图标覆盖时，Shell 扩展可能会调用 Initialize 来传递上下文信息。
    //   (4) 拖放操作（Drag-and-Drop Handlers）
    //   当用户拖放文件时，Initialize 方法也可能被调用，以初始化与拖放操作相关的数据。
    //
    //
    //[ComImport]
    //    这是一个 特性（Attribute），表示这个接口是从 COM 导入的，定义了一个 COM 接口。
    //    使用 ComImport 告诉 .NET，这个接口在 COM 中已经定义，我们只是通过该声明使用它。
    //
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //    指定了接口的类型。
    //    ComInterfaceType.InterfaceIsIUnknown 表示这个接口使用 IUnknown 类型的接口（COM 中的基本接口）。这意味着该接口遵循 IUnknown 的规范，主要包括：
    //    引用计数（AddRef 和 Release）。
    //    接口查询机制（QueryInterface）。
    //    IUnknown 是 COM 中所有接口的基础接口。
    //
    //[Guid("000214e8-0000-0000-c000-000000000046")]
    //   这个 GUID 是 IShellExtInit 在 COM 中的标识，Windows 系统使用它来找到该接口的实现
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214e8-0000-0000-c000-000000000046")]
    public interface IShellExtInit
    {
        //参数解析：
        //IntPtr pidlFolder:
        //   表示扩展所操作的文件夹的标识符（PIDL）。
        //   PIDL（Pointer to an Item Identifier List）是 Windows Shell 使用的一种结构，用来标识文件系统中的对象。
        //IntPtr dataObject:
        //   表示扩展所处理的数据对象（IDataObject 的指针）。
        //   数据对象通常包含关于文件或对象的信息（如文件名列表），供扩展程序使用。
        //IntPtr hKeyProgID:
        //   表示注册表中的 ProgID（程序标识符）的句柄，用于与特定的程序关联。
        //   在 IContextMenu.cs 中添加 IShellExtInit 接口定义
        void Initialize(IntPtr pidlFolder, IntPtr dataObject, IntPtr hKeyProgID);
    }
} 
