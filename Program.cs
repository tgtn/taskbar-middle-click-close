using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace TaskbarMiddleClickClose;

internal static class Program
{
    private const int WhMouseLl = 14;
    private const int WmMButtonDown = 0x0207;
    private const uint WmClose = 0x0010;

    private static readonly LowLevelMouseProc Proc = HookCallback;
    private static IntPtr _hookId;

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(Point point);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int X;
        public int Y;
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

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [STAThread]
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

        UnhookWindowsHookEx(_hookId);
    }

    private static IntPtr SetHook(
        LowLevelMouseProc proc)
    {
        using var process = Process.GetCurrentProcess();
        using var module = process.MainModule;

        return SetWindowsHookEx(
            WhMouseLl,
            proc,
            GetModuleHandle(module?.ModuleName),
            0);
    }

    private static IntPtr HookCallback(
        int nCode,
        IntPtr wParam,
        IntPtr lParam)
    {
        if (nCode >= 0 &&
            wParam == WmMButtonDown)
        {
            var hook =
                Marshal.PtrToStructure<MsllHookStruct>(
                    lParam);
            var hwnd = WindowFromPoint(hook.pt);

            if (hwnd != IntPtr.Zero &&
                IsTaskbarArea(hwnd))
            {
                QueueUiaWork(
                    hook.pt.X,
                    hook.pt.Y);

                return 1; // only block taskbar middle-click
            }
        }

        return CallNextHookEx(
            _hookId,
            nCode,
            wParam,
            lParam);
    }

    private static bool CloseTaskbarButtonAtPoint(
        int x,
        int y)
    {
        try
        {
            var button =
                GetTaskbarButtonAtPoint(x, y);

            if (button == null)
                return false;

            var taskbarTitle =
                NormalizeTaskbarTitle(
                    button.Current.Name);

            if (string.IsNullOrWhiteSpace(taskbarTitle))
                return false;

            Console.WriteLine(
                $"Clicked: {taskbarTitle}");

            var result =
                FindBestMatchingWindow(
                    taskbarTitle);

            Console.WriteLine(
                $"Best score: {result.Score}");

            if (result.Score < 25)
            {
                Console.WriteLine(
                    "Match score too low. Not closing anything.");

                return false;
            }

            SendMessage(
                result.Hwnd,
                WmClose,
                IntPtr.Zero,
                IntPtr.Zero);

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            return false;
        }
    }

    private static AutomationElement?
        GetTaskbarButtonAtPoint(
            int x,
            int y)
    {
        var element =
            AutomationElement.FromPoint(
                new System.Windows.Point(x, y));

        while (element != null)
        {
            if (Equals(element.Current.ControlType, ControlType.Button))
            {
                return element;
            }

            element =
                TreeWalker.RawViewWalker
                    .GetParent(element);
        }

        return null;
    }

    private static string NormalizeTaskbarTitle(
        string title)
    {
        if (string.IsNullOrWhiteSpace(title))
            return string.Empty;

        var idx =
            title.LastIndexOf(
                " - ",
                StringComparison.Ordinal);

        if (idx > 0)
        {
            var suffix = title[(idx + 3)..];

            if (suffix.Contains(
                    "running window",
                    StringComparison.OrdinalIgnoreCase))
            {
                title = title[..idx];
            }
        }

        return title.Trim();
    }

    private static (IntPtr Hwnd, int Score) FindBestMatchingWindow(
        string taskbarTitle)
    {
        var bestHwnd = IntPtr.Zero;
        var bestScore = 0;

        EnumWindows((hwnd, _) =>
        {
            if (!IsWindowVisible(hwnd))
                return true;

            var len = GetWindowTextLength(hwnd);

            if (len <= 0)
                return true;

            var sb = new StringBuilder(len + 1);

            GetWindowText(hwnd, sb, sb.Capacity);

            var title = sb.ToString();

            var score = CalculateScore(
                taskbarTitle,
                title);

            if (score > bestScore)
            {
                bestScore = score;
                bestHwnd = hwnd;
            }

            return true;
        }, IntPtr.Zero);

        return (bestHwnd, bestScore);
    }

    private static int CalculateScore(
        string taskbarTitle,
        string windowTitle)
    {
        taskbarTitle = taskbarTitle.ToLowerInvariant();
        windowTitle = windowTitle.ToLowerInvariant();

        if (windowTitle == taskbarTitle)
            return 1000;

        if (windowTitle.Contains(taskbarTitle))
            return 500 + taskbarTitle.Length;

        if (taskbarTitle.Contains(windowTitle))
            return 300 + windowTitle.Length;

        var words = taskbarTitle.Split(' ');

        return (from word in words where word.Length >= 4 where windowTitle.Contains(word) select word.Length).Sum();
    }

    private static void QueueUiaWork(
        int x,
        int y)
    {
        var thread = new Thread(() =>
        {
            try
            {
                var closed = CloseTaskbarButtonAtPoint(x, y);
                Console.WriteLine($"Windows closed: {closed}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        });

        thread.SetApartmentState(
            ApartmentState.STA);

        thread.Start();
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