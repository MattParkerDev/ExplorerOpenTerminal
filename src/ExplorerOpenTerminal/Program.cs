﻿using System.Diagnostics;
using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.UI.Input.KeyboardAndMouse;
using Windows.Win32.UI.Shell;
using Ardalis.GuardClauses;
using HWND = Windows.Win32.Foundation.HWND;
using IServiceProvider = Windows.Win32.System.Com.IServiceProvider;

namespace ExplorerOpenTerminal;

public static class Program
{
	private static readonly Guid _iidIShellBrowser = typeof(IShellBrowser).GUID;


	[STAThread]
	public static void Main()
	{
		var (activeWindow, className) = GetForegroundWindow();

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
		var activeWindow = PInvoke.GetForegroundWindow();
		var className = GetWindowClassName(activeWindow);
		return (activeWindow, className);
	}

	private static string GetWindowClassName(HWND windowHandle)
	{
		Span<char> charSpan = stackalloc char[256];
		PInvoke.GetClassName(windowHandle, charSpan);

		var className = charSpan.ToString().Trim('\0'); // Trim null characters
		return className;
	}

	// I don't like doing this, if there's a better way to get the previously active window, I'd prefer to do that
	private static void CallAltTab()
	{
		const int delay = 2;
		Thread.Sleep(delay);
		var altDownInput = new KEYBDINPUT { wVk = VIRTUAL_KEY.VK_MENU, wScan = 0xb8, dwFlags = 0, dwExtraInfo = 0 };
		var tabDownInput = new KEYBDINPUT { wVk = VIRTUAL_KEY.VK_TAB, wScan = 0x8f, dwFlags = 0, dwExtraInfo = 0 };
		var tabUpInput = tabDownInput with { dwFlags = KEYBD_EVENT_FLAGS.KEYEVENTF_KEYUP };
		var altUpInput = altDownInput with { dwFlags = KEYBD_EVENT_FLAGS.KEYEVENTF_KEYUP };

		// The user has most likely invoked this program via a shortcut on the desktop, along with a keybind. The shortcut must contain the alt key, and if still held, will mess with our alt tab
		// Lets wait until the alt key is released (and ctrl & tab lol)
		var altKeyState = PInvoke.GetAsyncKeyState((int)VIRTUAL_KEY.VK_MENU);
		var ctrlKeyState = PInvoke.GetAsyncKeyState((int)VIRTUAL_KEY.VK_CONTROL);
		var tabKeyState = PInvoke.GetAsyncKeyState((int)VIRTUAL_KEY.VK_TAB);
		while (altKeyState is not 0 || ctrlKeyState is not 0 || tabKeyState is not 0)
		{
			Thread.Sleep(20);
			altKeyState = PInvoke.GetAsyncKeyState((int)VIRTUAL_KEY.VK_MENU);
			ctrlKeyState = PInvoke.GetAsyncKeyState((int)VIRTUAL_KEY.VK_CONTROL);
			tabKeyState = PInvoke.GetAsyncKeyState((int)VIRTUAL_KEY.VK_TAB);
		}


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
		var activeTab = PInvoke.FindWindowEx(new HWND(hwnd), new HWND(IntPtr.Zero), "ShellTabWindowClass", null);

		var shellWindows = (IShellWindows)new ShellWindows();
		Guard.Against.Null(shellWindows);
		for (var i = 0; i < shellWindows.Count; i++)
		{
			var windowObject = shellWindows.Item(i);
			dynamic window = windowObject;
			var windowPointer = (IntPtr)window.HWND;
			if (window == null || windowPointer != hwnd)
				continue;

			var windowPath = window!.Document.Folder.Self.Path as string;
			var webBrowser2 = (IWebBrowser2) windowObject;

			var serviceProvider = (IServiceProvider) webBrowser2;

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
