using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaskbarMiddleClickClose;

internal static class TaskbarMiddleClickClose
{
    private const int WhMouseLl = 14;
    private const int WmMbuttondown = 0x0207;
    private const int WmMbuttonup = 0x0208;
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
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, Input[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern IntPtr GetTopWindow(IntPtr hWnd);

    private const uint GwHwndnext = 2;

    private const uint WmClose = 0x0010;

    private const uint InputMouse = 0;
    private const uint MouseEventfLeftDown = 0x0002;
    private const uint MouseEventfLeftUp = 0x0004;

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Msllhookstruct
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
        [FieldOffset(0)] public MouseInput mi;
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

        // Load icon from embedded resources
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
        aboutItem.Click += (_, _) =>
        {
            MessageBox.Show(
                "Taskbar Middle-Click Close\n\n" +
                "Middle-click any taskbar button to close that application.\n\n" +
                "Running in system tray.",
                "About",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        };

        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (_, _) =>
        {
            trayIcon.Visible = false;
            Application.Exit();
        };

        contextMenu.Items.Add(aboutItem);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(exitItem);

        trayIcon.ContextMenuStrip = contextMenu;

        // Handle application exit
        Application.ApplicationExit += (_, _) =>
        {
            UnhookWindowsHookEx(_hookId);
            trayIcon.Dispose();
        };

        Application.Run();
    }

    private static IntPtr SetHook(LowLevelMouseProc proc)
    {
        using var curProcess = Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule;
        return SetWindowsHookEx(WhMouseLl, proc, GetModuleHandle(curModule?.ModuleName), 0);
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            var hookStruct = Marshal.PtrToStructure<Msllhookstruct>(lParam);
            var pt = hookStruct.pt;

            switch (wParam)
            {
                case WmMbuttondown:
                {
                    var hwnd = WindowFromPoint(pt);
                    if (hwnd != IntPtr.Zero && IsTaskbarArea(hwnd))
                    {
                        Console.WriteLine($"Middle click on taskbar at ({pt.x}, {pt.y})");
                        _processingMiddleClick = true;

                        // Get snapshot of windows and Z-order before the click
                        var beforeWindow = GetForegroundWindow();
                        var zOrderBefore = GetWindowZOrder();

                        // Unhook temporarily to avoid catching our own simulated click
                        UnhookWindowsHookEx(_hookId);

                        // Simulate left click at the exact position
                        var inputs = new Input[2];

                        inputs[0].type = InputMouse;
                        inputs[0].u.mi.dwFlags = MouseEventfLeftDown;

                        inputs[1].type = InputMouse;
                        inputs[1].u.mi.dwFlags = MouseEventfLeftUp;

                        _ = SendInput(2, inputs, Marshal.SizeOf(typeof(Input)));

                        // Re-establish hook
                        using (var curProcess = Process.GetCurrentProcess())
                        using (var curModule = curProcess.MainModule)
                        {
                            _hookId = SetWindowsHookEx(WhMouseLl, Proc, GetModuleHandle(curModule?.ModuleName), 0);
                        }

                        // Wait and close the window that got activated
                        Task.Run(async () =>
                        {
                            await Task.Delay(150);

                            var targetWindow = IntPtr.Zero;
                            var afterWindow = GetForegroundWindow();
                            var zOrderAfter = GetWindowZOrder();

                            // Strategy 1: Check if foreground window changed
                            if (afterWindow != IntPtr.Zero && afterWindow != beforeWindow && IsWindow(afterWindow))
                            {
                                if (IsValidTargetWindow(afterWindow))
                                {
                                    targetWindow = afterWindow;
                                    Console.WriteLine("Strategy 1: Foreground window changed");
                                }
                            }

                            // Strategy 2: Check Z-order changes (window brought to front)
                            if (targetWindow == IntPtr.Zero && zOrderAfter.Count > 0)
                            {
                                // Find window that moved to the top of Z-order
                                var topWindow = zOrderAfter[0];
                                if (IsValidTargetWindow(topWindow) &&
                                    (zOrderBefore.Count == 0 || zOrderBefore[0] != topWindow))
                                {
                                    targetWindow = topWindow;
                                    Console.WriteLine("Strategy 2: Z-order changed");
                                }
                            }

                            // Strategy 3: Check for same window that was already foreground (clicked current window)
                            if (targetWindow == IntPtr.Zero && afterWindow == beforeWindow && IsValidTargetWindow(afterWindow))
                            {
                                targetWindow = afterWindow;
                                Console.WriteLine("Strategy 3: Same window (already focused)");
                            }

                            if (targetWindow != IntPtr.Zero)
                            {
                                _ = GetWindowThreadProcessId(targetWindow, out var processId);

                                if (processId != 0)
                                {
                                    try
                                    {
                                        var proc = Process.GetProcessById((int)processId);
                                        Console.WriteLine($"Closing: {proc.ProcessName}");
                                        SendMessage(targetWindow, WmClose, IntPtr.Zero, IntPtr.Zero);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error: {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Could not determine target window");
                            }

                            _processingMiddleClick = false;
                        });

                        // Block the middle click
                        return 1;
                    }

                    break;
                }
                case WmMbuttonup when _processingMiddleClick:
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

    private static bool IsValidTargetWindow(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
            return false;

        var className = new StringBuilder(256);
        _ = GetClassName(hWnd, className, className.Capacity);
        var cls = className.ToString();

        // Exclude system windows
        return !cls.Contains("Shell_TrayWnd") && !cls.Contains("Shell_SecondaryTrayWnd") &&
               !cls.Contains("Progman") && !cls.Contains("WorkerW") &&
               !cls.Contains("Windows.UI.Core.CoreWindow");
    }

    private static List<IntPtr> GetWindowZOrder()
    {
        var windows = new List<IntPtr>();
        var hWnd = GetTopWindow(IntPtr.Zero);

        while (hWnd != IntPtr.Zero)
        {
            if (IsWindowVisible(hWnd) && GetParent(hWnd) == IntPtr.Zero)
            {
                var length = GetWindowTextLength(hWnd);
                if (length > 0 && IsValidTargetWindow(hWnd))
                {
                    windows.Add(hWnd);
                }
            }

            hWnd = GetWindow(hWnd, GwHwndnext);
        }

        return windows;
    }
}