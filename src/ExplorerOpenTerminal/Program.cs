using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Text;
using System.Windows.Automation;
using Windows.Win32;
using Windows.Win32.UI.Shell;
using Vanara.Extensions;
using Vanara.Extensions.Reflection;
using Vanara.PInvoke;
using Vanara.Windows.Shell;
using HWND = Windows.Win32.Foundation.HWND;
using IServiceProvider = Windows.Win32.System.Com.IServiceProvider;

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

    [STAThread]
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
	    GC.Collect();
	    GC.WaitForPendingFinalizers(); // if we try to create another Shell.Application it will blow up unless the COM object has been released
	    //var shellWindows123 = (IShellWindows)new ShellWindows();
	    var path = handleAndFolderPaths.First(s => s.Handle == activeWindow);
	    GetActiveExplorerTab2(path.Handle);
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
	    //Marshal.ReleaseComObject(shellApp);
	    //Marshal.FinalReleaseComObject(shellApp);
	    //GC.Collect();
	    //GC.WaitForPendingFinalizers();
	    return handleAndFolderPaths;
	}

	static string? GetActiveExplorerTab(IntPtr hwnd)
	{
		// Try to retrieve the active tab
		var activeTabHandle = User32.FindWindowEx(hwnd, IntPtr.Zero, "ShellTabWindowClass");
		var activeTabPointer = activeTabHandle.DangerousGetHandle();
		if (activeTabPointer == IntPtr.Zero)
			return null;

		// Use Shell.Application to get the correct tab
		var shellAppType = Type.GetTypeFromProgID("Shell.Application");
		dynamic? shellApp = Activator.CreateInstance(shellAppType);
		if (shellApp == null)
			return null;

		var windowTextSpan = new Span<char>(new char[256]);
		PInvoke.GetWindowText(new HWND(activeTabPointer), windowTextSpan);
		var windowText = windowTextSpan.ToString();


		var windows = shellApp!.Windows();
		foreach (var window in windows)
		{
			var windowPointer = new IntPtr((int)window.HWND);
			if (window == null || windowPointer != hwnd)
				continue;

			if (window is not object windowObject)
				continue;

			var folderPath = window.Document.Folder.Self.Path as string;

			windowObject.QueryInterface(Shell32.SID_SShellBrowser, out var shellBrowserasdf);

			//var explorerBrowser = (Shell32.IExplorerBrowser)Marshal.GetObjectForIUnknown(window);
			//var explorerTest = Marshal.GetComInterfaceForObject<Shell32.IExplorerBrowser, Shell32.IExplorerBrowser>(window);
			// Query IShellBrowser
			var shellBrowserPtr = Marshal.GetIUnknownForObject(windowObject);
			Guid IID_IShellBrowser = new("000214E2-0000-0000-C000-000000000046");
			Guid IID_IShellBrowser2 = Shell32.SID_SShellBrowser;

			IntPtr shellBrowser;
			Marshal.QueryInterface(shellBrowserPtr, in Shell32.SID_SInternetExplorer, out shellBrowser);
			uint thisTab = 0;
			var ptr = Marshal.ReadIntPtr(shellBrowser, 3 * IntPtr.Size); // Call VTable index 3 for active tab
			//DllCall(ptr, out thisTab);

			if (thisTab == (uint)activeTabPointer)
			{
				return window.Document.Folder.Self.Path as string;
			}
		}
		return null;
	}

	static dynamic GetActiveExplorerTab2(IntPtr hwnd)
	{
		IntPtr activeTab = PInvoke.FindWindowEx(new HWND(hwnd), new HWND(IntPtr.Zero), "ShellTabWindowClass", null); // Windows 11 File Explorer
		if (activeTab == IntPtr.Zero)
		{
			activeTab = PInvoke.FindWindowEx(new HWND(hwnd), new HWND(IntPtr.Zero), "TabWindowClass1", null); // Internet Explorer
		}

		var vanaraShellWindows = (Shell32.IShellWindows)new Shell32.ShellWindows();
		for (var i = 0; i < vanaraShellWindows.Count; i++)
		{
			var windowObject = vanaraShellWindows.Item(i);
			dynamic window = windowObject;
			var cast = (IWebBrowser) windowObject;
			var asdfasdf = new object();
			windowObject.QueryInterface(typeof(IServiceProvider).GUID, out var serviceProviderObject);
			var serviceProvider3 = (IServiceProvider) windowObject;
			Guid SID_STopLevelBrowser = new Guid("4C96BE40-915C-11CF-99D3-00AA004AE837");
			Guid IID_IShellBrowser = typeof(IShellBrowser).GUID;
			serviceProvider3.QueryService(in SID_STopLevelBrowser, in IID_IShellBrowser, out var ppv);
			var serviceProvider2 = (IServiceProvider) serviceProviderObject;
			serviceProvider2.QueryService(PInvoke.SID_STopLevelBrowser, typeof(IShellBrowser).GUID, out var shellBrowserObject2);
			var shellBrowser2 = (IShellBrowser) shellBrowserObject2;
			var testsetase = cast.GetPropertyValue<long>("HWND");
			var test2 = window.Document.Folder.Self.Path as string;
		}
		IShellWindows shellWindows123 = (IShellWindows)new ShellWindows();
		for (var i = 0; i < shellWindows123.Count; i++)
		{
			var window = shellWindows123.Item(i);
		}

		var serviceProvider = (IServiceProvider)shellWindows123.FindWindowSW(
			PInvoke.CSIDL_DESKTOP,
			pvarLocRoot: null,
			ShellWindowTypeConstants.SWC_DESKTOP,
			phwnd: out _,
			ShellWindowFindWindowOptions.SWFO_NEEDDISPATCH);

		serviceProvider.QueryService(PInvoke.SID_STopLevelBrowser, typeof(IShellBrowser).GUID, out var shellBrowserObject);
		var shellBrowser123 = (IShellBrowser)shellBrowserObject;

		shellBrowser123.QueryActiveShellView(out var shellView);

		dynamic shellApp = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
		dynamic shellWindows = shellApp.Windows();
		foreach (var window in shellWindows)
		{
			var folderPath = window.Document.Folder.Self.Path as string;
			if ((IntPtr)window.HWND != hwnd)
				continue;

			if (activeTab != IntPtr.Zero)
			{
				const string IID_IShellBrowserString = "000214E2-0000-0000-C000-000000000046";
				Guid IID_IShellBrowser = new Guid(IID_IShellBrowserString);
				IntPtr shellBrowser;

				var test = typeof(IShellBrowser).GUID;
				IntPtr pUnknown = Marshal.GetIUnknownForObject(window);

				if (pUnknown == IntPtr.Zero)
					continue; // Skip if the COM object is invalid
				int hr = Marshal.QueryInterface(pUnknown, ref IID_IShellBrowser, out shellBrowser);
				if (hr != 0)
					continue;

				int thisTab = 0;
				Marshal.WriteInt32(shellBrowser, 3, thisTab);

				if ((IntPtr)thisTab != activeTab)
					continue;
			}
			return window;
		}
		return null;
	}

	private static void DllCall(IntPtr shellBrowser, out uint thisTab)
	{
		unsafe
		{
			IntPtr vTable = Marshal.ReadIntPtr(shellBrowser);
			IntPtr getActiveTabPtr = Marshal.ReadIntPtr(vTable, 3 * IntPtr.Size);
			var getActiveTab = (delegate* unmanaged[Stdcall]<IntPtr, out uint, int>)getActiveTabPtr;
			//var thisTabPointer = IntPtr.Zero;
			getActiveTab(shellBrowser, out thisTab);
			//thisTab = 0;
		}
	}
}
