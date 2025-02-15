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

    static void Main()
    {
		// Get previous focus before terminal opens
        //IntPtr activeWindow = GetPreviouslyFocusedWindow();
        var activeWindow = Vanara.PInvoke.User32.GetForegroundWindow();
        //IntPtr activeWindow = GetForegroundWindow();

        // Get class name of the active window
        var classNameBuilder = new StringBuilder(256);
        User32.GetClassName(activeWindow, classNameBuilder, classNameBuilder.Capacity);

        var className = classNameBuilder.ToString();
        // Check if the active window is a Windows Explorer instance
        if (className is "CabinetWClass")
        {
            User32.GetWindowThreadProcessId(activeWindow, out var processId);
            if (processId == 0)
            {
                Console.WriteLine("Failed to get process ID.");
                return;
            }

            var process = Process.GetProcessById((int)processId);
            if (process?.ProcessName is not "explorer")
            {
                Console.WriteLine("Active window is not an Explorer window.");
                return;
            }

            //var automationElement = AutomationElement.FromHandle(activeWindow.DangerousGetHandle());
            //ArgumentNullException.ThrowIfNull(automationElement);
            //
            //var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
            //AutomationElement editElement = automationElement.FindFirst(TreeScope.Descendants, condition);
            //ArgumentNullException.ThrowIfNull(editElement);
            //string folderPath = editElement?.GetCurrentPropertyValue(ValuePattern.ValueProperty) as string;
            var folderPath = GetActiveExplorerPath();
            Console.WriteLine("Active Explorer Path: " + folderPath);
        }
        else
        {
            Console.WriteLine("Active window is not an Explorer window.");
        }

        Console.ReadLine();
    }

    static string GetActiveExplorerPath()
    {
        try
        {
            // Create Shell.Application COM object
            Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
            dynamic shellApp = Activator.CreateInstance(shellAppType);

            // Get all open Explorer windows
            dynamic windows = shellApp.Windows();

            foreach (var window in windows)
            {
                if (window == null) continue;

                IntPtr hwnd = new IntPtr((int)window.HWND);

                // Compare hwnd with the active window
                if (hwnd == GetForegroundWindow())
                {
                    return window.Document.Folder.Self.Path;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }

        return null;
    }
}
