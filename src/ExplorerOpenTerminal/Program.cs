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

    static string? GetActiveExplorerPath(IntPtr activeWindow)
    {
	    var pathFromActiveTab = GetActiveExplorerTabPath(activeWindow);
	    GC.Collect();
	    GC.WaitForPendingFinalizers(); // if we try to create another Shell.Application it will blow up unless the COM object has been released
	    return pathFromActiveTab;
    }

	static string? GetActiveExplorerTabPath(IntPtr hwnd)
	{
		IntPtr activeTab = PInvoke.FindWindowEx(new HWND(hwnd), new HWND(IntPtr.Zero), "ShellTabWindowClass", null); // Windows 11 File Explorer

		IShellWindows shellWindows123 = (IShellWindows)new ShellWindows();
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
			var serviceProvider = (IServiceProvider) serviceProviderObject;
			Guid SID_STopLevelBrowser = new Guid("4C96BE40-915C-11CF-99D3-00AA004AE837");
			Guid IID_IShellBrowser = typeof(IShellBrowser).GUID;
			serviceProvider.QueryService(in SID_STopLevelBrowser, in IID_IShellBrowser, out var ppv);
			var shellBrowser2 = (IShellBrowser) ppv;

			IntPtr comPtr = Marshal.GetComInterfaceForObject(shellBrowser2, typeof(IShellBrowser));
			IntPtr vTable = Marshal.ReadIntPtr(comPtr);
			int start = Marshal.GetStartComSlot(typeof(IShellBrowser));
			int end = Marshal.GetEndComSlot(typeof(IShellBrowser));

			ComMemberType mType = 0;
			for (var j = start; j <= end; j++)
			{
				//System.Reflection.MemberInfo mi = Marshal.GetMethodInfoForComSlot(typeof(IShellBrowser), i, ref mType);
				var functionPointer2 = Marshal.ReadIntPtr(vTable, j * Marshal.SizeOf<nint>());
				Console.WriteLine("Method {0} at address 0x{1:X}", "mi.Name", functionPointer2.ToInt64());
			}
			var functionPointer = Marshal.ReadIntPtr(vTable, start * Marshal.SizeOf<nint>());
			var delegate2 = Marshal.GetDelegateForFunctionPointer<GetActiveTabDelegate>(functionPointer);
			delegate2(comPtr, out var thisTab123);
			var supposedActiveTabPointer = (IntPtr) thisTab123;
			if (supposedActiveTabPointer == activeTab)
			{
				return windowPath;
			}
		}

		return null;
	}

	[UnmanagedFunctionPointer(CallingConvention.StdCall)]
	delegate int GetActiveTabDelegate(IntPtr shellBrowser, out uint thisTabPointer);
}
