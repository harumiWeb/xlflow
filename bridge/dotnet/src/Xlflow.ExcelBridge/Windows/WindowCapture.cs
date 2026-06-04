using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

namespace Xlflow.ExcelBridge.Windows;

internal sealed record WindowInfo(long Hwnd, string Title, string ClassName, int Left, int Top, int Width, int Height);

internal sealed record WindowImageInfo(int WidthPx, int HeightPx, double Scale);

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Window capture is Windows-only bridge behavior.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Window capture falls back from optional DPI and screen-copy APIs.")]
[SuppressMessage("Performance", "CA1838:Avoid StringBuilder parameters for P/Invokes", Justification = "The bounded text helpers mirror the existing DialogWatcher implementation.")]
[SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance", Justification = "The helper returns iteration-friendly collections for simple call sites.")]
internal static class WindowCapture
{
    private const int SwRestore = 9;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpNoActivate = 0x0010;
    private const uint MonitorDefaultToNearest = 2;

    public static WindowInfo? GetWindowInfo(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindowVisible(hwnd))
        {
            return null;
        }

        if (!NativeMethods.GetWindowRect(hwnd, out var rect))
        {
            return null;
        }

        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0)
        {
            return null;
        }

        return new WindowInfo(
            hwnd.ToInt64(),
            NativeMethods.GetWindowText(hwnd),
            NativeMethods.GetClassName(hwnd),
            rect.Left,
            rect.Top,
            width,
            height);
    }

    public static WindowInfo? WaitForStableWindow(IntPtr hwnd, TimeSpan timeout, TimeSpan pollInterval, int stableSamples)
    {
        var deadline = DateTime.UtcNow + timeout;
        var lastSignature = "";
        var stableCount = 0;
        WindowInfo? last = null;
        while (DateTime.UtcNow < deadline)
        {
            var window = GetWindowInfo(hwnd);
            if (window is null)
            {
                Thread.Sleep(pollInterval);
                continue;
            }

            var signature = $"{window.Left}:{window.Top}:{window.Width}:{window.Height}";
            if (string.Equals(signature, lastSignature, StringComparison.Ordinal))
            {
                stableCount++;
            }
            else
            {
                lastSignature = signature;
                stableCount = 1;
            }

            last = window;
            if (stableCount >= stableSamples)
            {
                return window;
            }
            Thread.Sleep(pollInterval);
        }

        return last;
    }

    public static bool IsLikelyUserFormWindow(WindowInfo? window)
    {
        if (window is null)
        {
            return false;
        }

        if (string.Equals(window.ClassName.Trim(), "XLMAIN", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return window.ClassName.StartsWith("Thunder", StringComparison.OrdinalIgnoreCase);
    }

    public static WindowInfo? FindWindowByTitle(int processId, string title, bool exactMatch)
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            return null;
        }

        foreach (var hwnd in EnumerateTopLevelWindows(processId))
        {
            var info = GetWindowInfo(hwnd);
            if (info is null)
            {
                continue;
            }

            if (exactMatch)
            {
                if (string.Equals(info.Title, title, StringComparison.Ordinal))
                {
                    return info;
                }
                continue;
            }

            if (info.Title.Contains(title, StringComparison.OrdinalIgnoreCase))
            {
                return info;
            }
        }

        return null;
    }

    public static WindowInfo? MoveWindowIntoCaptureBounds(WindowInfo? window)
    {
        if (window is null)
        {
            return null;
        }

        var hwnd = new IntPtr(window.Hwnd);
        if (NativeMethods.IsIconic(hwnd))
        {
            _ = NativeMethods.ShowWindow(hwnd, SwRestore);
        }
        _ = NativeMethods.SetForegroundWindow(hwnd);

        var targetRect = GetWorkingArea(hwnd);
        if (targetRect is null)
        {
            return WaitForStableWindow(hwnd, TimeSpan.FromMilliseconds(1200), TimeSpan.FromMilliseconds(100), 2) ?? window;
        }

        var margin = 16;
        var minLeft = targetRect.Value.Left + margin;
        var maxLeft = targetRect.Value.Right - window.Width - margin;
        if (maxLeft < minLeft)
        {
            minLeft = targetRect.Value.Left;
            maxLeft = Math.Max(minLeft, targetRect.Value.Right - window.Width);
        }

        var minTop = targetRect.Value.Top + margin;
        var maxTop = targetRect.Value.Bottom - window.Height - margin;
        if (maxTop < minTop)
        {
            minTop = targetRect.Value.Top;
            maxTop = Math.Max(minTop, targetRect.Value.Bottom - window.Height);
        }

        var left = Math.Min(Math.Max(window.Left, minLeft), maxLeft);
        var top = Math.Min(Math.Max(window.Top, minTop), maxTop);
        _ = NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, left, top, 0, 0, SwpNoSize | SwpNoZOrder | SwpNoActivate);
        return WaitForStableWindow(hwnd, TimeSpan.FromMilliseconds(1200), TimeSpan.FromMilliseconds(100), 2)
            ?? GetWindowInfo(hwnd)
            ?? window;
    }

    public static WindowImageInfo CaptureWindowImage(long hwndValue, string path)
    {
        var hwnd = new IntPtr(hwndValue);
        if (!NativeMethods.GetWindowRect(hwnd, out var rect))
        {
            throw new InvalidOperationException("window_not_found: failed to resolve window bounds");
        }

        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0)
        {
            throw new InvalidOperationException("image_capture_failed: target window has non-positive bounds");
        }

        var scale = GetWindowCaptureScale(hwnd);
        var captureWidth = (int)Math.Ceiling(width * scale);
        var captureHeight = (int)Math.Ceiling(height * scale);

        using var bitmap = new Bitmap(captureWidth, captureHeight);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);

        var printOk = false;
        var hdc = IntPtr.Zero;
        try
        {
            hdc = graphics.GetHdc();
            printOk = NativeMethods.PrintWindow(hwnd, hdc, 2) || NativeMethods.PrintWindow(hwnd, hdc, 0);
        }
        finally
        {
            if (hdc != IntPtr.Zero)
            {
                graphics.ReleaseHdc(hdc);
            }
        }

        if (!printOk)
        {
            try
            {
                graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(captureWidth, captureHeight));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"image_capture_failed: failed to capture the target window ({ex.Message})", ex);
            }
        }

        using var trimmed = TrimBlackEdges(bitmap);
        trimmed.Save(path, ImageFormat.Png);
        return new WindowImageInfo(trimmed.Width, trimmed.Height, scale);
    }

    private static List<IntPtr> EnumerateTopLevelWindows(int processId)
    {
        var handles = new List<IntPtr>();
        NativeMethods.EnumWindows((hwnd, lParam) =>
        {
            _ = NativeMethods.GetWindowThreadProcessId(hwnd, out var pid);
            if (processId <= 0 || pid == processId)
            {
                handles.Add(hwnd);
            }
            return true;
        }, IntPtr.Zero);
        return handles;
    }

    private static NativeMethods.Rect? GetWorkingArea(IntPtr hwnd)
    {
        var monitor = NativeMethods.MonitorFromWindow(hwnd, MonitorDefaultToNearest);
        if (monitor == IntPtr.Zero)
        {
            return null;
        }

        var info = new NativeMethods.MonitorInfo();
        info.CbSize = Marshal.SizeOf<NativeMethods.MonitorInfo>();
        if (!NativeMethods.GetMonitorInfo(monitor, ref info))
        {
            return null;
        }

        return info.RcWork;
    }

    private static double GetWindowCaptureScale(IntPtr hwnd)
    {
        try
        {
            var dpi = NativeMethods.GetDpiForWindow(hwnd);
            if (dpi < 96)
            {
                return 1.0;
            }
            return Math.Clamp(dpi / 96.0, 1.0, 4.0);
        }
        catch
        {
            return 1.0;
        }
    }

    private static Bitmap TrimBlackEdges(Bitmap bitmap)
    {
        var left = 0;
        var right = bitmap.Width - 1;
        var top = 0;
        var bottom = bitmap.Height - 1;

        while (left < right && EdgeIsBlack(bitmap, "left", left))
        {
            left++;
        }
        while (right > left && EdgeIsBlack(bitmap, "right", right))
        {
            right--;
        }
        while (top < bottom && EdgeIsBlack(bitmap, "top", top))
        {
            top++;
        }
        while (bottom > top && EdgeIsBlack(bitmap, "bottom", bottom))
        {
            bottom--;
        }

        var width = right - left + 1;
        var height = bottom - top + 1;
        if (width <= 0 || height <= 0 || (width == bitmap.Width && height == bitmap.Height))
        {
            return new Bitmap(bitmap);
        }

        var trimmed = new Bitmap(width, height, bitmap.PixelFormat);
        using var graphics = Graphics.FromImage(trimmed);
        graphics.DrawImage(bitmap, new Rectangle(0, 0, width, height), new Rectangle(left, top, width, height), GraphicsUnit.Pixel);
        return trimmed;
    }

    private static bool EdgeIsBlack(Bitmap bitmap, string edge, int index, int threshold = 64, int step = 2, double minimumDarkRatio = 0.75, int ignoreTopPixels = 32)
    {
        var samples = 0;
        var darkSamples = 0;
        switch (edge)
        {
            case "left":
            case "right":
                for (var y = Math.Min(ignoreTopPixels, Math.Max(0, bitmap.Height - 1)); y < bitmap.Height; y += step)
                {
                    var pixel = bitmap.GetPixel(index, y);
                    samples++;
                    if (pixel.R <= threshold && pixel.G <= threshold && pixel.B <= threshold)
                    {
                        darkSamples++;
                    }
                }
                break;
            case "top":
            case "bottom":
                for (var x = 0; x < bitmap.Width; x += step)
                {
                    var pixel = bitmap.GetPixel(x, index);
                    samples++;
                    if (pixel.R <= threshold && pixel.G <= threshold && pixel.B <= threshold)
                    {
                        darkSamples++;
                    }
                }
                break;
        }

        return samples > 0 && (double)darkSamples / samples >= minimumDarkRatio;
    }

    private static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MonitorInfo
        {
            public int CbSize;
            public Rect RcMonitor;
            public Rect RcWork;
            public uint DwFlags;
        }

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hwnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassNameW(IntPtr hwnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint flags);

        [DllImport("user32.dll")]
        public static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        public static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo monitorInfo);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int cx, int cy, uint flags);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hwnd, int command);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hwnd);

        public static string GetWindowText(IntPtr hwnd)
        {
            var text = new StringBuilder(1024);
            _ = GetWindowTextW(hwnd, text, text.Capacity);
            return text.ToString();
        }

        public static string GetClassName(IntPtr hwnd)
        {
            var text = new StringBuilder(256);
            _ = GetClassNameW(hwnd, text, text.Capacity);
            return text.ToString();
        }
    }
}
