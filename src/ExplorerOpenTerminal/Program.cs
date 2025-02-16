using System.Diagnostics;
using System.Text;
using Windows.Win32;
using Windows.Win32.System.Variant;
using Windows.Win32.UI.Input.KeyboardAndMouse;
using Windows.Win32.UI.Shell;
using Ardalis.GuardClauses;
using Vanara.Extensions;
using Vanara.PInvoke;
using HWND = Windows.Win32.Foundation.HWND;
using IServiceProvider = Windows.Win32.System.Com.IServiceProvider;

namespace ExplorerOpenTerminal;

public class Program
{
	private static readonly Guid _iidIShellBrowser = typeof(IShellBrowser).GUID;


	[STAThread]
	public static void Main()
	{
		// do not start the shortcut to this as minimised - minimised windows are not allowed to set focus to other windows
		//Thread.Sleep(500);
		var (activeWindow, className) = GetFocusedWindow();
		if (className is not "CabinetWClass")
		{
			Console.WriteLine($"Active window ({className}) is not an Explorer window, calling Alt+Tab");
			CallAltTab();
			(activeWindow, className) = GetFocusedWindow();
		}
		if (className is not "CabinetWClass")
		{
			Console.WriteLine($"Active window ({className}) is not an Explorer window.");
			return;
		}

		var folderPath = GetActiveExplorerPath(activeWindow);
		Console.WriteLine("Active Explorer Path: " + folderPath);
		if (string.IsNullOrWhiteSpace(folderPath))
		{
			Console.WriteLine("Active explorer path is null, returning");
			return;
		}
		LaunchWindowsTerminalInDirectory(folderPath);
	}

	private static (IntPtr, string) GetFocusedWindow()
	{
		var activeWindow = User32.GetForegroundWindow().DangerousGetHandle();
		var className = GetWindowName(activeWindow);
		return (activeWindow, className);
	}

	private static string GetWindowName(IntPtr windowHandle)
	{
		// Get class name of the active window
		var classNameBuilder = new StringBuilder(256);
		User32.GetClassName(windowHandle, classNameBuilder, classNameBuilder.Capacity);

		var className = classNameBuilder.ToString();
		return className;
	}

	private static void CallAltTab()
	{
		const int delay = 10;
		Thread.Sleep(delay);
		PInvoke.keybd_event((byte)VIRTUAL_KEY.VK_MENU,0xb8,0 , 0); //Alt Press
		Thread.Sleep(delay);
		PInvoke.keybd_event((byte)VIRTUAL_KEY.VK_TAB,0x8f,0 , 0); // Tab Press
		Thread.Sleep(delay);
		PInvoke.keybd_event((byte)VIRTUAL_KEY.VK_TAB,0x8f, KEYBD_EVENT_FLAGS.KEYEVENTF_KEYUP,0); // Tab Release
		Thread.Sleep(delay);
		PInvoke.keybd_event((byte)VIRTUAL_KEY.VK_MENU,0xb8,KEYBD_EVENT_FLAGS.KEYEVENTF_KEYUP,0); // Alt Release
		Thread.Sleep(delay);
		//SendKeys.SendWait("%{Tab}");
		//SendKeys.Flush();
		//Thread.Sleep(10);
	}

	private static void LaunchWindowsTerminalInDirectory(string directory)
	{
		var process = new Process
		{
			StartInfo = new ProcessStartInfo
			{
				FileName = "wt",
				Arguments = $"""-d "{directory}" """,
				UseShellExecute = false,
				CreateNoWindow = true,
			}
		};
		process.Start();
	}

	private static string? GetActiveExplorerPath(IntPtr activeWindow)
	{
		var pathFromActiveTab = GetActiveExplorerTabPath(activeWindow);
		GC.Collect();
		GC.WaitForPendingFinalizers(); // if we try to create another Shell.Application it will blow up unless the COM object has been released // Not necessary in this case, but keeping for reference
		return pathFromActiveTab;
	}

	private static string? GetActiveExplorerTabPath(IntPtr hwnd)
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

			shellBrowser.GetWindow(out var windowHandle);
			if (windowHandle == activeTab)
			{
				return windowPath;
			}

		}

		return null;
	}
}
