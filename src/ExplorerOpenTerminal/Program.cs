using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using Vanara.PInvoke;

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
		// Get previous focus before terminal opens
        //IntPtr activeWindow = GetPreviouslyFocusedWindow();
        //Thread.Sleep(1000);
        var activeWindow = User32.GetForegroundWindow().DangerousGetHandle();
        IntPtr activeWindow2 = GetForegroundWindow();

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
	    return "";
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

	    foreach (var window in windows)
	    {
		    if (window == null) continue;

		    var handlerPointer = new IntPtr((int)window.HWND);
		    var folderPath = window.Document.Folder.Self.Path as string;

		    handleAndFolderPaths.Add(new HandleAndFolderPath(handlerPointer, folderPath));
	    }
	    return handleAndFolderPaths;
	}
}
