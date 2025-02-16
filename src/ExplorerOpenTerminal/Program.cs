using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Windows.Win32;
using Windows.Win32.UI.Input.KeyboardAndMouse;
using Windows.Win32.UI.Shell;
using Ardalis.GuardClauses;
using Vanara.Extensions;
using Vanara.PInvoke;
using HWND = Windows.Win32.Foundation.HWND;
using IServiceProvider = Windows.Win32.System.Com.IServiceProvider;

namespace ExplorerOpenTerminal;

public static class Program
{
	private static readonly Guid _iidIShellBrowser = typeof(IShellBrowser).GUID;


	[STAThread]
	public static void Main()
	{
		// if this program is started via a shortcut, the window that is focused is Shell_TrayWnd, which is the taskbar
		// Shell_TrayWnd does not like calling alt tab, so we just wait a bit
		// Waiting however, does not change what the foreground window is, it just prevents alt tab getting stuck

		var (activeWindow, className) = GetForegroundWindow();
		if (className is "Shell_TrayWnd")
		{
			Thread.Sleep(100);
			(activeWindow, className) = GetForegroundWindow();
		}

		if (className is not "CabinetWClass")
		{
			Console.WriteLine($"Active window ({className}) is not an Explorer window, calling Alt+Tab");
			CallAltTab();
			(activeWindow, className) = GetForegroundWindow();
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

	private static (IntPtr, string) GetForegroundWindow()
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
		const int delay = 2;
		Thread.Sleep(delay);
		var altDownInput = new KEYBDINPUT { wVk = VIRTUAL_KEY.VK_MENU, wScan = 0xb8, dwFlags = 0, dwExtraInfo = 0 };
		var tabDownInput = new KEYBDINPUT { wVk = VIRTUAL_KEY.VK_TAB, wScan = 0x8f, dwFlags = 0, dwExtraInfo = 0 };
		var tabUpInput = tabDownInput with { dwFlags = KEYBD_EVENT_FLAGS.KEYEVENTF_KEYUP };
		var altUpInput = altDownInput with { dwFlags = KEYBD_EVENT_FLAGS.KEYEVENTF_KEYUP };

		PInvoke.SendInput([new INPUT {type = INPUT_TYPE.INPUT_KEYBOARD, Anonymous = { ki = altDownInput}}], Marshal.SizeOf<INPUT>());
		Thread.Sleep(delay);
		PInvoke.SendInput([new INPUT {type = INPUT_TYPE.INPUT_KEYBOARD, Anonymous = { ki = tabDownInput}}], Marshal.SizeOf<INPUT>());
		Thread.Sleep(delay);
		PInvoke.SendInput([new INPUT {type = INPUT_TYPE.INPUT_KEYBOARD, Anonymous = { ki = tabUpInput}}], Marshal.SizeOf<INPUT>());
		Thread.Sleep(delay);
		PInvoke.SendInput([new INPUT {type = INPUT_TYPE.INPUT_KEYBOARD, Anonymous = { ki = altUpInput}}], Marshal.SizeOf<INPUT>());
		Thread.Sleep(delay);
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
				CreateNoWindow = false,
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
		IntPtr activeTab = PInvoke.FindWindowEx(new HWND(hwnd), new HWND(IntPtr.Zero), "ShellTabWindowClass", null);

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
