using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using Application = System.Windows.Forms.Application;
using MessageBox = System.Windows.Forms.MessageBox;
using Point = System.Windows.Point;

namespace TaskbarMiddleClickClose;

internal static class Program
{
    private const int WhMouseLl = 14;
    private const int WmMButtonDown = 0x0207;
    private const uint WmClose = 0x0010;

    private static readonly LowLevelMouseProc Proc = HookCallback;

    private static IntPtr _hookId;
    private static bool _hookRemoved;

    [DllImport(
        "user32.dll",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern IntPtr FindWindow(
        string? lpClassName,
        string? lpWindowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(
        EnumWindowsProc lpEnumFunc,
        IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(
        IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(
        IntPtr hWnd,
        out NativeRect lpRect);

    [DllImport(
        "user32.dll",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern int GetWindowText(
        IntPtr hWnd,
        StringBuilder lpString,
        int nMaxCount);

    [DllImport(
        "user32.dll",
        SetLastError = true)]
    private static extern int GetWindowTextLength(
        IntPtr hWnd);

    [DllImport(
        "user32.dll",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern int GetClassName(
        IntPtr hWnd,
        StringBuilder lpClassName,
        int nMaxCount);

    [DllImport(
        "user32.dll",
        SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(
        IntPtr hWnd,
        uint msg,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport(
        "user32.dll",
        SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(
        int idHook,
        LowLevelMouseProc lpfn,
        IntPtr hMod,
        uint dwThreadId);

    [DllImport(
        "user32.dll",
        SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(
        IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(
        IntPtr hhk,
        int nCode,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport(
        "kernel32.dll",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern IntPtr GetModuleHandle(
        string? lpModuleName);

    private delegate IntPtr LowLevelMouseProc(
        int nCode,
        IntPtr wParam,
        IntPtr lParam);

    private delegate bool EnumWindowsProc(
        IntPtr hWnd,
        IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public readonly bool Contains(NativePoint point)
        {
            return point.X >= Left &&
                   point.X < Right &&
                   point.Y >= Top &&
                   point.Y < Bottom;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MsllHookStruct
    {
        public NativePoint Point;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        _hookId = SetHook(Proc);

        if (_hookId == IntPtr.Zero)
        {
            var error = Marshal.GetLastWin32Error();

            MessageBox.Show(
                $"The mouse hook could not be installed.\n\n" +
                $"Windows error: {error}",
                "Taskbar Middle-Click Close",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            return;
        }

        using var trayIcon = CreateTrayIcon();

        Application.ApplicationExit += (_, _) =>
        {
            RemoveMouseHook();
            trayIcon.Visible = false;
        };

        trayIcon.Visible = true;

        try
        {
            Application.Run();
        }
        finally
        {
            trayIcon.Visible = false;
            RemoveMouseHook();
        }
    }

    private static NotifyIcon CreateTrayIcon()
    {
        var trayIcon = new NotifyIcon
        {
            Text = "Taskbar Middle-Click Close",
            Icon = LoadApplicationIcon()
        };

        var contextMenu = new ContextMenuStrip();

        var aboutItem = new ToolStripMenuItem("About");

        aboutItem.Click += (_, _) =>
        {
            MessageBox.Show(
                "Taskbar Middle-Click Close\n\n" +
                "Middle-click a running application's taskbar " +
                "button to close that application's window.\n\n" +
                "Running in the system tray.",
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

        return trayIcon;
    }

    private static Icon LoadApplicationIcon()
    {
        try
        {
            var iconPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Assets",
                "app-icon.ico");

            if (File.Exists(iconPath))
                return new Icon(iconPath);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(
                $"Could not load app-icon.ico: {ex}");
        }

        return SystemIcons.Application;
    }

    private static IntPtr SetHook(
        LowLevelMouseProc proc)
    {
        using var process = Process.GetCurrentProcess();
        using var module = process.MainModule;

        var moduleHandle = GetModuleHandle(
            module?.ModuleName);

        return SetWindowsHookEx(
            WhMouseLl,
            proc,
            moduleHandle,
            0);
    }

    private static void RemoveMouseHook()
    {
        if (_hookRemoved || _hookId == IntPtr.Zero)
            return;

        _hookRemoved = true;

        UnhookWindowsHookEx(_hookId);
        _hookId = IntPtr.Zero;
    }

    private static IntPtr HookCallback(
        int nCode,
        IntPtr wParam,
        IntPtr lParam)
    {
        if (nCode < 0 ||
            wParam != new IntPtr(WmMButtonDown))
        {
            return CallNextHookEx(
                _hookId,
                nCode,
                wParam,
                lParam);
        }

        try
        {
            var hook =
                Marshal.PtrToStructure<MsllHookStruct>(
                    lParam);

            if (!IsPointInsideTaskbar(hook.Point))
            {
                return CallNextHookEx(
                    _hookId,
                    nCode,
                    wParam,
                    lParam);
            }

            QueueUiaWork(
                hook.Point.X,
                hook.Point.Y);

            // Suppress the normal Windows taskbar middle-click action,
            // which would otherwise open another application instance.
            return new IntPtr(1);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(
                $"Mouse hook error: {ex}");

            return CallNextHookEx(
                _hookId,
                nCode,
                wParam,
                lParam);
        }
    }

    private static bool IsPointInsideTaskbar(
        NativePoint point)
    {
        var primaryTaskbar = FindWindow(
            "Shell_TrayWnd",
            null);

        if (IsPointInsideWindow(
                primaryTaskbar,
                point))
        {
            return true;
        }

        var insideSecondaryTaskbar = false;

        EnumWindows((hWnd, _) =>
        {
            if (!string.Equals(
                    GetWindowClassName(hWnd),
                    "Shell_SecondaryTrayWnd",
                    StringComparison.Ordinal))
            {
                return true;
            }

            if (!IsPointInsideWindow(hWnd, point))
                return true;

            insideSecondaryTaskbar = true;

            // Stop enumerating.
            return false;
        }, IntPtr.Zero);

        return insideSecondaryTaskbar;
    }

    private static bool IsPointInsideWindow(
        IntPtr hWnd,
        NativePoint point)
    {
        if (hWnd == IntPtr.Zero)
            return false;

        return GetWindowRect(
                   hWnd,
                   out var rect) &&
               rect.Contains(point);
    }

    private static string GetWindowClassName(
        IntPtr hWnd)
    {
        var className = new StringBuilder(256);

        var length = GetClassName(
            hWnd,
            className,
            className.Capacity);

        return length > 0
            ? className.ToString()
            : string.Empty;
    }

    private static void QueueUiaWork(
        int x,
        int y)
    {
        var thread = new Thread(() =>
        {
            try
            {
                var closed = CloseTaskbarButtonAtPoint(
                    x,
                    y);

                Debug.WriteLine(
                    $"Window close requested: {closed}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"UI Automation worker error: {ex}");
            }
        })
        {
            IsBackground = true,
            Name = "TaskbarMiddleClickClose.UIA"
        };

        thread.SetApartmentState(
            ApartmentState.STA);

        thread.Start();
    }

    private static bool CloseTaskbarButtonAtPoint(
        int x,
        int y)
    {
        try
        {
            var button = GetTaskbarButtonAtPoint(
                x,
                y);

            if (button == null)
            {
                Debug.WriteLine(
                    "No taskbar application button found.");

                return false;
            }

            var elementName =
                button.Current.Name ?? string.Empty;

            var automationId =
                button.Current.AutomationId ?? string.Empty;

            var className =
                button.Current.ClassName ?? string.Empty;

            Debug.WriteLine(
                $"UIA name: '{elementName}'");

            Debug.WriteLine(
                $"UIA automation ID: '{automationId}'");

            Debug.WriteLine(
                $"UIA class: '{className}'");

            var taskbarTitle =
                NormalizeTaskbarTitle(elementName);

            if (string.IsNullOrWhiteSpace(taskbarTitle))
            {
                Debug.WriteLine(
                    "The taskbar button has no usable name.");

                return false;
            }

            Debug.WriteLine(
                $"Normalized title: '{taskbarTitle}'");

            var result =
                FindBestMatchingWindow(taskbarTitle);

            Debug.WriteLine(
                $"Matched title: '{result.Title}'");

            Debug.WriteLine(
                $"Match score: {result.Score}");

            if (result.Hwnd == IntPtr.Zero ||
                result.Score < 25)
            {
                Debug.WriteLine(
                    "No sufficiently reliable window match found.");

                return false;
            }

            if (!PostMessage(
                    result.Hwnd,
                    WmClose,
                    IntPtr.Zero,
                    IntPtr.Zero))
            {
                Debug.WriteLine(
                    $"PostMessage failed: " +
                    $"{Marshal.GetLastWin32Error()}");

                return false;
            }

            return true;
        }
        catch (ElementNotAvailableException ex)
        {
            Debug.WriteLine(
                $"Taskbar element disappeared: {ex}");

            return false;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(
                $"Close operation failed: {ex}");

            return false;
        }
    }

    private static AutomationElement? GetTaskbarButtonAtPoint(
        int x,
        int y)
    {
        AutomationElement root;

        try
        {
            root = AutomationElement.FromPoint(
                new Point(x, y));
        }
        catch (ElementNotAvailableException)
        {
            return null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }

        return FindTaskbarButtonAtPoint(
            root,
            x,
            y,
            0);
    }

    private static AutomationElement? FindTaskbarButtonAtPoint(
        AutomationElement parent,
        int x,
        int y,
        int depth)
    {
        if (depth > 20)
            return null;

        AutomationElement? child;

        try
        {
            child = TreeWalker.RawViewWalker.GetFirstChild(
                parent);
        }
        catch (ElementNotAvailableException)
        {
            return null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }

        while (child != null)
        {
            AutomationElement? nextSibling = null;

            try
            {
                nextSibling =
                    TreeWalker.RawViewWalker.GetNextSibling(
                        child);

                var rectangle =
                    child.Current.BoundingRectangle;

                if (!rectangle.IsEmpty &&
                    IsPointInsideRectangle(
                        rectangle,
                        x,
                        y))
                {
                    if (IsTaskbarApplicationButton(child))
                        return child;

                    var descendant =
                        FindTaskbarButtonAtPoint(
                            child,
                            x,
                            y,
                            depth + 1);

                    if (descendant != null)
                        return descendant;
                }
            }
            catch (ElementNotAvailableException)
            {
                // The taskbar tree can change while traversing it.
            }
            catch (InvalidOperationException)
            {
                // Ignore inaccessible UI Automation elements.
            }

            child = nextSibling;
        }

        return null;
    }

    private static bool IsPointInsideRectangle(
        Rect rectangle,
        int x,
        int y)
    {
        return x >= rectangle.Left &&
               x < rectangle.Right &&
               y >= rectangle.Top &&
               y < rectangle.Bottom;
    }

    private static bool IsTaskbarApplicationButton(
        AutomationElement element)
    {
        try
        {
            var className =
                element.Current.ClassName ?? string.Empty;

            if (!className.Equals(
                    "Taskbar.TaskListButtonAutomationPeer",
                    StringComparison.Ordinal))
            {
                return false;
            }

            var controlType =
                element.Current.ControlType;

            if (!Equals(controlType, ControlType.Button))
                return false;

            var automationId =
                element.Current.AutomationId ??
                string.Empty;

            if (!automationId.StartsWith(
                    "Window:",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var name =
                element.Current.Name ??
                string.Empty;

            return !string.IsNullOrWhiteSpace(name);
        }
        catch (ElementNotAvailableException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    private static string NormalizeTaskbarTitle(
        string title)
    {
        if (string.IsNullOrWhiteSpace(title))
            return string.Empty;

        title = title.Trim();

        var separatorIndex = title.LastIndexOf(
            " - ",
            StringComparison.Ordinal);

        if (separatorIndex > 0)
        {
            var suffix = title[(separatorIndex + 3)..];

            if (suffix.Contains(
                    "running window",
                    StringComparison.OrdinalIgnoreCase) ||
                suffix.Equals(
                    "running",
                    StringComparison.OrdinalIgnoreCase))
            {
                title = title[..separatorIndex];
            }
        }

        string[] statusSuffixes =
        [
            " - active",
            " - inactive",
            " - attention",
            " - flashing"
        ];

        foreach (var suffix in statusSuffixes)
        {
            if (!title.EndsWith(
                    suffix,
                    StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            title = title[..^suffix.Length];
            break;
        }

        return title.Trim();
    }

    private static WindowMatch FindBestMatchingWindow(
        string taskbarTitle)
    {
        var bestMatch = new WindowMatch(
            IntPtr.Zero,
            string.Empty,
            0);

        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd))
                return true;

            var title = GetNativeWindowTitle(hWnd);

            if (string.IsNullOrWhiteSpace(title))
                return true;

            var score = CalculateScore(
                taskbarTitle,
                title);

            if (score <= bestMatch.Score)
                return true;

            bestMatch = new WindowMatch(
                hWnd,
                title,
                score);

            return true;
        }, IntPtr.Zero);

        return bestMatch;
    }

    private static string GetNativeWindowTitle(
        IntPtr hWnd)
    {
        var length = GetWindowTextLength(hWnd);

        if (length <= 0)
            return string.Empty;

        var title = new StringBuilder(length + 1);

        var copiedLength = GetWindowText(
            hWnd,
            title,
            title.Capacity);

        return copiedLength > 0
            ? title.ToString()
            : string.Empty;
    }

    private static int CalculateScore(
        string taskbarTitle,
        string windowTitle)
    {
        taskbarTitle = NormalizeForComparison(
            taskbarTitle);

        windowTitle = NormalizeForComparison(
            windowTitle);

        if (taskbarTitle.Length == 0 ||
            windowTitle.Length == 0)
        {
            return 0;
        }

        if (windowTitle.Equals(
                taskbarTitle,
                StringComparison.Ordinal))
        {
            return 1000;
        }

        if (windowTitle.Contains(
                taskbarTitle,
                StringComparison.Ordinal))
        {
            return 500 + taskbarTitle.Length;
        }

        if (taskbarTitle.Contains(
                windowTitle,
                StringComparison.Ordinal))
        {
            return 300 + windowTitle.Length;
        }

        var words = taskbarTitle.Split(
            ' ',
            StringSplitOptions.RemoveEmptyEntries |
            StringSplitOptions.TrimEntries);

        var score = 0;

        foreach (var word in words)
        {
            if (word.Length < 4)
                continue;

            if (windowTitle.Contains(
                    word,
                    StringComparison.Ordinal))
            {
                score += word.Length;
            }
        }

        return score;
    }

    private static string NormalizeForComparison(
        string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var builder = new StringBuilder(
            value.Length);

        var previousWasSpace = false;

        foreach (var character in value)
        {
            if (char.IsWhiteSpace(character))
            {
                if (!previousWasSpace)
                {
                    builder.Append(' ');
                    previousWasSpace = true;
                }

                continue;
            }

            builder.Append(
                char.ToLowerInvariant(character));

            previousWasSpace = false;
        }

        return builder.ToString().Trim();
    }

    private readonly record struct WindowMatch(
        IntPtr Hwnd,
        string Title,
        int Score);
}