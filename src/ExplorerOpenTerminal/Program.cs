using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using Vanara.PInvoke;
using Vanara.Windows.Shell;

class Program
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

    [StructLayout(LayoutKind.Sequential)]
    struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int left, top, right, bottom;
    }

    public record HandleAndFolderPath(IntPtr Handle, string? FolderPath);

    public static void Main()
    {
	    var activeWindow = User32.GetForegroundWindow().DangerousGetHandle();

        // Get class name of the active window
        var classNameBuilder = new StringBuilder(256);
        User32.GetClassName(activeWindow, classNameBuilder, classNameBuilder.Capacity);

        var className = classNameBuilder.ToString();
        // Check if the active window is a Windows Explorer instance
        if (className is not "CabinetWClass")
        {
	        Console.WriteLine("Active window is not an Explorer window.");
	        return;
        }

        var folderPath = GetActiveExplorerPath(activeWindow);
        Console.WriteLine("Active Explorer Path: " + folderPath);

        Console.ReadLine();
    }

    static string GetActiveExplorerPath(IntPtr activeWindow)
    {
	    var handleAndFolderPaths = GetExplorerWindows();
	    var path = handleAndFolderPaths.First(s => s.Handle == activeWindow);
	    GetActiveExplorerTab(path.Handle);
	    return path.FolderPath;
    }

    private static List<HandleAndFolderPath> GetExplorerWindows()
	{
	    List<HandleAndFolderPath> handleAndFolderPaths = [];

	    // Create Shell.Application COM object
	    var shellAppType = Type.GetTypeFromProgID("Shell.Application");
	    ArgumentNullException.ThrowIfNull(shellAppType);
	    dynamic? shellApp = Activator.CreateInstance(shellAppType);
	    ArgumentNullException.ThrowIfNull(shellApp);
	    //Shell32.IShellWindows shellWindows = (Shell32.IShellWindows)shellApp;
	    Vanara.Windows.Shell.ShellBrowser shellBrowser2 = new ShellBrowser();
	    var test2 = shellBrowser2.Text;
	    var items = shellBrowser2.Items;


	    //Shell32.IShellWindows shellWindows2 = new ;

	    // Get all open Explorer windows
	    var windows = shellApp!.Windows();

	    foreach (var window in windows) // Window is apparently an IWebBrowser2
	    {
		    if (window == null) continue;

		    var handlerPointer = new IntPtr((int)window.HWND);
		    var folderPath = window.Document.Folder.Self.Path as string;
		    var test = window.LocationName;
		    var test3 = window.LocationURL;

		    handleAndFolderPaths.Add(new HandleAndFolderPath(handlerPointer, folderPath));
	    }
	    return handleAndFolderPaths;
	}

	static string? GetActiveExplorerTab(IntPtr hwnd)
	{
		// Try to retrieve the active tab
		IntPtr activeTab = IntPtr.Zero;

		var activeTabHandle = User32.FindWindowEx(hwnd, IntPtr.Zero, "ShellTabWindowClass");
		var activeTabPointer = activeTabHandle.DangerousGetHandle();
		if (activeTabPointer == IntPtr.Zero)
			return null;

		// Use Shell.Application to get the correct tab
		var shellAppType = Type.GetTypeFromProgID("Shell.Application");
		dynamic? shellApp = Activator.CreateInstance(shellAppType);
		if (shellApp == null)
			return null;

		var windows = shellApp!.Windows();
		foreach (var window in windows)
		{
			if (window == null || new IntPtr((int)window.HWND) != hwnd)
				continue;

			// Query IShellBrowser
			var shellBrowserPtr = Marshal.GetIUnknownForObject(window);
			Guid IID_IShellBrowser = new("000214E2-0000-0000-C000-000000000046");
			IntPtr shellBrowser;
			Marshal.QueryInterface(shellBrowserPtr, ref IID_IShellBrowser, out shellBrowser);


			uint thisTab;
			Marshal.ReadIntPtr(shellBrowser, 3 * IntPtr.Size); // Call VTable index 3 for active tab
			DllCall(shellBrowser, out thisTab);

			if (thisTab == (uint)activeTab)
			{
				return window.Document.Folder.Self.Path as string;
			}
		}
		return null;
	}

	private static void DllCall(IntPtr shellBrowser, out uint thisTab)
	{
		IntPtr vTable = Marshal.ReadIntPtr(shellBrowser);
		IntPtr getActiveTabPtr = Marshal.ReadIntPtr(vTable, 3 * IntPtr.Size);
		//var getActiveTab = (delegate* unmanaged[Stdcall]<IntPtr, out uint, int>)getActiveTabPtr;
		//getActiveTab(shellBrowser, out thisTab);
		thisTab = 0;
	}
}
