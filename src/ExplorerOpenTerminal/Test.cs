using System;
using Windows.Win32.UI.Shell;
using Windows.Win32;
using System.Runtime.InteropServices;
using Windows.Win32.Foundation;
using Windows.Win32.UI.WindowsAndMessaging;

internal static class ProcessUtils
{
    public static void StartProcessWithoutElevation(string fileName, string workingDirectory)
    {
        var shellWindows = (IShellWindows)new ShellWindows();

        var serviceProvider = (IServiceProvider)shellWindows.FindWindowSW(
            PInvoke.CSIDL_DESKTOP,
            pvarLocRoot: null,
            ShellWindowTypeConstants.SWC_DESKTOP,
            phwnd: out _,
            ShellWindowFindWindowOptions.SWFO_NEEDDISPATCH);

        var shellBrowser = (IShellBrowser)serviceProvider.QueryService(PInvoke.SID_STopLevelBrowser, typeof(IShellBrowser).GUID);

        shellBrowser.QueryActiveShellView(out var shellView);

        shellView.GetItemObject((uint)_SVGIO.SVGIO_BACKGROUND, typeof(IDispatch).GUID, out var folderViewAsObject);

        var folderView = (IShellFolderViewDual)folderViewAsObject;
        var shellDispatch = (IShellDispatch2)folderView.Application;

        var fileNameAsBstr = (BSTR)Marshal.StringToBSTR(fileName);
        try
        {
            shellDispatch.ShellExecute(File: fileNameAsBstr, vArgs: null, vDir: workingDirectory, vOperation: "", SHOW_WINDOW_CMD.SW_NORMAL);
        }
        finally
        {
            Marshal.FreeBSTR(fileNameAsBstr);
        }
    }

    // Workaround for https://github.com/microsoft/CsWin32/issues/860
    [Guid("85CB6900-4D95-11CF-960C-0080C7F4EE85"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), ComImport]
    private interface IShellWindows
    {
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object FindWindowSW([MarshalAs(UnmanagedType.Struct)] in object pvarLoc, [MarshalAs(UnmanagedType.Struct)] in object? pvarLocRoot, ShellWindowTypeConstants swClass, out int phwnd, ShellWindowFindWindowOptions swfwOptions);
    }

    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), ComImport]
    private interface IServiceProvider
    {
        [return: MarshalAs(UnmanagedType.Interface)]
        object QueryService(in Guid guidService, in Guid riid);
    }

    // Workaround for https://github.com/microsoft/CsWin32/issues/861
    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    private interface IDispatch
    {
    }
}
