using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace TaskbarMiddleClickClose;

class TaskbarMiddleClickClose
{
    private const int WH_MOUSE_LL = 14;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_MBUTTONUP = 0x0208;
    private static readonly LowLevelMouseProc Proc = HookCallback;
    private static IntPtr _hookId = IntPtr.Zero;
    private static bool _processingMiddleClick;

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(Point point);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, Input[] pInputs, int cbSize);

    private const uint WM_CLOSE = 0x0010;
    
    private const uint INPUT_MOUSE = 0;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MsllHookStruct
    {
        public Point pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public uint type;
        public InputUnion u;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MouseInput mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        
        _hookId = SetHook(Proc);
        
        // Create system tray icon
        var trayIcon = new NotifyIcon();
        trayIcon.Text = "Taskbar Middle-Click Close";
        
        // Try to load custom icon, fall back to default if not found
        try
        {
            var iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app-icon.ico");
            trayIcon.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;
        }
        catch
        {
            trayIcon.Icon = SystemIcons.Application;
        }
        
        trayIcon.Visible = true;
        
        // Create context menu
        var contextMenu = new ContextMenuStrip();
        
        var aboutItem = new ToolStripMenuItem("About");
        aboutItem.Click += (_, _) => {
            MessageBox.Show(
                "Taskbar Middle-Click Close\n\n" +
                "Middle-click any taskbar button to close that application.\n\n" +
                "Running in system tray.",
                "About",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        };
        
        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (_, _) => {
            trayIcon.Visible = false;
            Application.Exit();
        };
        
        contextMenu.Items.Add(aboutItem);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(exitItem);
        
        trayIcon.ContextMenuStrip = contextMenu;
        
        // Handle application exit
        Application.ApplicationExit += (_, _) => {
            UnhookWindowsHookEx(_hookId);
            trayIcon.Dispose();
        };
        
        Application.Run();
    }

    private static IntPtr SetHook(LowLevelMouseProc proc)
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;
        return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule?.ModuleName), 0);
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            var hookStruct = Marshal.PtrToStructure<MsllHookStruct>(lParam);
            var pt = hookStruct.pt;
            
            switch (wParam)
            {
                case WM_MBUTTONDOWN:
                {
                    var hwnd = WindowFromPoint(pt);
                    if (hwnd != IntPtr.Zero && IsTaskbarArea(hwnd))
                    {
                        Console.WriteLine($"Middle click on taskbar at ({pt.x}, {pt.y})");
                        _processingMiddleClick = true;
                    
                        // Get current foreground window before the click
                        var beforeWindow = GetForegroundWindow();
                    
                        // Unhook temporarily to avoid catching our own simulated click
                        UnhookWindowsHookEx(_hookId);
                    
                        // Simulate left click at the exact position
                        var inputs = new Input[2];
                    
                        inputs[0].type = INPUT_MOUSE;
                        inputs[0].u.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
                    
                        inputs[1].type = INPUT_MOUSE;
                        inputs[1].u.mi.dwFlags = MOUSEEVENTF_LEFTUP;
                    
                        _ = SendInput(2, inputs, Marshal.SizeOf<Input>());
                    
                        // Re-establish hook
                        using (var curProcess = Process.GetCurrentProcess())
                        using (var curModule = curProcess.MainModule)
                        {
                            _hookId = SetWindowsHookEx(WH_MOUSE_LL, Proc, GetModuleHandle(curModule?.ModuleName), 0);
                        }
                    
                        // Wait and close the window that got activated
                        Task.Run(async () =>
                        {
                            await Task.Delay(250);
                        
                            var afterWindow = GetForegroundWindow();
                        
                            // Check if a different window got focus
                            if (afterWindow != IntPtr.Zero && afterWindow != beforeWindow && IsWindow(afterWindow))
                            {
                                _ = GetWindowThreadProcessId(afterWindow, out var processId);
                            
                                if (processId != 0)
                                {
                                    try
                                    {
                                        var proc = Process.GetProcessById((int)processId);
                                    
                                        var className = new StringBuilder(256);
                                        _ = GetClassName(afterWindow, className, className.Capacity);
                                        var cls = className.ToString();
                                    
                                        if (!cls.Contains("Shell_TrayWnd") && !cls.Contains("Progman") && 
                                            !cls.Contains("WorkerW"))
                                        {
                                            Console.WriteLine($"Closing: {proc.ProcessName}");
                                            SendMessage(afterWindow, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error: {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("No window change detected");
                            }
                        
                            _processingMiddleClick = false;
                        });
                    
                        // Block the middle click
                        return 1;
                    }

                    break;
                }
                case WM_MBUTTONUP when _processingMiddleClick:
                    // Block the middle button up event
                    return 1;
            }
        }
        
        return CallNextHookEx(_hookId, nCode, wParam, lParam);
    }

    private static bool IsTaskbarArea(IntPtr hwnd)
    {
        var checkWindow = hwnd;
        
        for (var i = 0; i < 10; i++)
        {
            var className = new StringBuilder(256);
            _ = GetClassName(checkWindow, className, className.Capacity);
            var cls = className.ToString();
            
            if (cls.Contains("Shell_TrayWnd") || 
                cls.Contains("Shell_SecondaryTrayWnd") ||
                cls.Contains("TaskList") || 
                cls.Contains("MSTaskSwWClass") ||
                cls.Contains("ReBarWindow32") ||
                cls.Contains("ToolbarWindow32"))
            {
                return true;
            }
            
            var parent = GetParent(checkWindow);
            if (parent == IntPtr.Zero) break;
            checkWindow = parent;
        }
        
        return false;
    }
}