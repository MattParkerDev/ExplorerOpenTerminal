﻿using System.Runtime.InteropServices;
using System.Text;
using Windows.Win32;
using Windows.Win32.UI.Shell;
using Ardalis.GuardClauses;
using Vanara.Extensions;
using Vanara.PInvoke;
using HWND = Windows.Win32.Foundation.HWND;
using IServiceProvider = Windows.Win32.System.Com.IServiceProvider;

class Program
{
	private static readonly Guid _iidIShellBrowser = typeof(IShellBrowser).GUID;

	[UnmanagedFunctionPointer(CallingConvention.StdCall)]
	private delegate int GetActiveTabDelegate(IntPtr shellBrowser, out IntPtr thisTabPointer);

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

        //Console.ReadLine();
    }

    static string? GetActiveExplorerPath(IntPtr activeWindow)
    {
	    var pathFromActiveTab = GetActiveExplorerTabPath(activeWindow);
	    GC.Collect();
	    GC.WaitForPendingFinalizers(); // if we try to create another Shell.Application it will blow up unless the COM object has been released
	    return pathFromActiveTab;
    }

	static string? GetActiveExplorerTabPath(IntPtr hwnd)
	{
		// The first result is the active tab
		IntPtr activeTab = PInvoke.FindWindowEx(new HWND(hwnd), new HWND(IntPtr.Zero), "ShellTabWindowClass", null); // Windows 11 File Explorer

		var shellWindows123 = (IShellWindows)new ShellWindows();
		Guard.Against.Null(shellWindows123);
		for (var i = 0; i < shellWindows123.Count; i++)
		{
			var windowObject = shellWindows123.Item(i);
			dynamic window = windowObject;
			var windowPointer = new IntPtr(window.HWND);
			if (window == null || windowPointer != hwnd)
				continue;

			var windowPath = window!.Document.Folder.Self.Path as string;
			var webBrowser2 = (IWebBrowser2) windowObject;
			webBrowser2.QueryInterface(typeof(IServiceProvider).GUID, out var serviceProviderObject);
			Guard.Against.Null(serviceProviderObject);
			var serviceProvider = (IServiceProvider) serviceProviderObject;

			serviceProvider.QueryService(in PInvoke.SID_STopLevelBrowser, in _iidIShellBrowser, out var shellBrowserObject);
			var shellBrowser = (IShellBrowser) shellBrowserObject;

			IntPtr comPtr = Marshal.GetComInterfaceForObject(shellBrowser, typeof(IShellBrowser));
			IntPtr vTable = Marshal.ReadIntPtr(comPtr);
			int start = Marshal.GetStartComSlot(typeof(IShellBrowser));
			int end = Marshal.GetEndComSlot(typeof(IShellBrowser));

			//ComMemberType mType = 0;
			for (var j = start; j <= end; j++)
			{
				//System.Reflection.MemberInfo memberInfo = Marshal.GetMethodInfoForComSlot(typeof(IShellBrowser), i, ref mType);
				var functionAddress = Marshal.ReadIntPtr(vTable, j * Marshal.SizeOf<nint>()).ToInt64();
				Console.WriteLine("Method {0} at address 0x{1:X}", "memberInfo.Name", functionAddress);
			}
			var functionPointer = Marshal.ReadIntPtr(vTable, start * Marshal.SizeOf<nint>());
			var delegate2 = Marshal.GetDelegateForFunctionPointer<GetActiveTabDelegate>(functionPointer);
			delegate2(comPtr, out var windowTabPointer);
			if (windowTabPointer == activeTab)
			{
				return windowPath;
			}
		}

		return null;
	}
}
