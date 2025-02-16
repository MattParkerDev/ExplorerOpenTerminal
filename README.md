# ExplorerOpenTerminal

This is a simple Windows program that allows you to open a Windows Terminal in the directory currently open in the focused Explorer window.

e.g. Ctrl+Alt+A
![explorerOpenTerminal](https://github.com/user-attachments/assets/84e96531-90f6-4e6f-a0d8-b74bf7058f36)

## How to use
1. Download the latest version from ![releases](https://github.com/MattParkerDev/ExplorerOpenTerminal/releases)
2. Save somewhere such as `C:/Users/Matt/Documents/ExplorerOpenTerminal/ExplorerOpenTerminal-v1.0.0.exe`
3. Add a shortcut to the exe on the desktop

	![image](https://github.com/user-attachments/assets/cb44c096-2e61-473c-bcf6-addd3ba1ab36)

4. Add a shortcut key to invoke the shortcut. I use Ctrl+Alt+A

	![image](https://github.com/user-attachments/assets/32701795-d8fe-42ef-beb6-858b0006df21)

5. Have an explorer open where you wish to open a terminal
6. Press your shortcut key!
<br>

Alternatively, the program can be invoked by pinning it to the task bar and invoking via Win+{{Index on the task bar}} e.g. Win+1

![explorerOpenTerminalpin](https://github.com/user-attachments/assets/e6026bbd-87de-4834-9fa1-4afadec09aa7)

## Notes
* The release binary only supports x64. You can compile for x86 if desired by updating `<PlatformTarget>x64</PlatformTarget>` in `ExplorerOpenTerminal.csproj` to x86 and running `publish.bat` 🙂
* The .NET 9 runtime is required to be installed, as the binary is published as `--self-contained false`
