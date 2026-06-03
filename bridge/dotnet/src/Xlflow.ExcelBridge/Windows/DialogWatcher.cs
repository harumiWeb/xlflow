using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;

namespace Xlflow.ExcelBridge.Windows;

public enum DialogKind
{
    Any,
    Runtime,
    Compile,
    MsgBox,
    InputBox,
    FileDialog,
}

public enum DialogActionPolicy
{
    ObserveOnly,
    SuppressVbaError,
    CancelSupportedNativeUi,
}

public sealed record DialogWatchRequest(
    int ExcelProcessId,
    long ExcelMainHwnd,
    DialogKind Kind,
    DialogActionPolicy ActionPolicy,
    TimeSpan Timeout,
    TimeSpan PollInterval,
    long VbeHwnd = 0,
    int VbeThreadId = 0);

public sealed record DialogElementSnapshot(
    long Hwnd,
    string ClassName,
    string Text,
    string? AutomationId = null,
    string? Name = null,
    string? ControlType = null);

public sealed record DialogSnapshot
{
    public string Kind { get; init; } = "";
    public long DetectedAtMs { get; init; }
    public string[] Sources { get; init; } = [];
    public long Hwnd { get; init; }
    public int Pid { get; init; }
    public int ThreadId { get; init; }
    public long OwnerHwnd { get; init; }
    public long RootOwnerHwnd { get; init; }
    public string Title { get; init; } = "";
    public string ClassName { get; init; } = "";
    public bool Visible { get; init; }
    public string ProcessImage { get; init; } = "";
    public string? AutomationId { get; init; }
    public string? Name { get; init; }
    public string? ControlType { get; init; }
    public string[] Text { get; init; } = [];
    public DialogElementSnapshot[] Buttons { get; init; } = [];
    public DialogElementSnapshot[] Children { get; init; } = [];
    public string Action { get; init; } = "";
    public string ActionMethod { get; init; } = "";
    public string ActionTarget { get; init; } = "";
    public bool ActionSucceeded { get; init; }
}

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Dialog watching is Windows-only bridge behavior.")]
public sealed class DialogWatcher
{
    private readonly IWindowEnumerator _windows;
    private readonly IUiaDialogAdapter _uia;

    public DialogWatcher()
        : this(new Win32WindowEnumerator(), new ComUiaDialogAdapter())
    {
    }

    internal DialogWatcher(IWindowEnumerator windows, IUiaDialogAdapter uia)
    {
        _windows = windows;
        _uia = uia;
    }

    public DialogSnapshot? WaitForDialog(DialogWatchRequest request, CancellationToken cancellationToken = default)
    {
        if (request.ExcelProcessId <= 0)
        {
            return null;
        }

        var stopwatch = Stopwatch.StartNew();
        var timeout = request.Timeout <= TimeSpan.Zero ? TimeSpan.FromSeconds(10) : request.Timeout;
        var pollInterval = request.PollInterval <= TimeSpan.Zero ? TimeSpan.FromMilliseconds(50) : request.PollInterval;

        while (stopwatch.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var candidate in _windows.Enumerate(request.ExcelProcessId, request.VbeThreadId))
            {
                if (!BelongsToTarget(candidate, request))
                {
                    continue;
                }

                var uia = _uia.Describe(candidate.Hwnd);
                var kind = DialogFingerprint.Classify(candidate, uia);
                if (kind is null || !Matches(request.Kind, kind.Value))
                {
                    continue;
                }

                var action = DialogActionSelector.Select(kind.Value, candidate, request.ActionPolicy);
                var actionResult = ExecuteAction(action, candidate, uia);
                return BuildSnapshot(candidate, uia, kind.Value, stopwatch.ElapsedMilliseconds, action, actionResult);
            }

            Thread.Sleep(pollInterval);
        }

        return null;
    }

    public IReadOnlyList<DialogSnapshot> CaptureCurrentDialogs(DialogWatchRequest request)
    {
        return CaptureCurrentDialogs(request, includeUia: true);
    }

    public IReadOnlyList<DialogSnapshot> CaptureCurrentDialogs(DialogWatchRequest request, bool includeUia)
    {
        var result = new List<DialogSnapshot>();
        foreach (var candidate in _windows.Enumerate(request.ExcelProcessId, request.VbeThreadId))
        {
            if (!BelongsToTarget(candidate, request))
            {
                continue;
            }
            var uia = includeUia ? _uia.Describe(candidate.Hwnd) : null;
            var kind = DialogFingerprint.Classify(candidate, uia);
            if (kind is null || !Matches(request.Kind, kind.Value))
            {
                continue;
            }
            result.Add(BuildSnapshot(candidate, uia, kind.Value, 0, DialogAction.None, DialogActionResult.None));
        }
        return result;
    }

    private DialogActionResult ExecuteAction(DialogAction action, WindowCandidate candidate, UiaDialogDescription? uia)
    {
        if (action == DialogAction.None)
        {
            return DialogActionResult.None;
        }

        if (action.TargetHwnd != 0 && _uia.TryInvoke(action.TargetHwnd))
        {
            return new DialogActionResult(true, "uia_invoke");
        }

        if (action.TargetHwnd != 0 && _windows.ClickButton(action.TargetHwnd))
        {
            return new DialogActionResult(true, "bm_click");
        }

        if (action.AllowWindowClose && _windows.CloseWindow(candidate.Hwnd))
        {
            return new DialogActionResult(true, "wm_close");
        }

        return new DialogActionResult(false, "");
    }

    private static DialogSnapshot BuildSnapshot(
        WindowCandidate candidate,
        UiaDialogDescription? uia,
        DialogKind kind,
        long detectedAtMs,
        DialogAction action,
        DialogActionResult actionResult)
    {
        var sources = uia is null ? new[] { "win32" } : new[] { "win32", "uia" };
        return new DialogSnapshot
        {
            Kind = kind.ToString().ToLowerInvariant(),
            DetectedAtMs = detectedAtMs,
            Sources = sources,
            Hwnd = candidate.Hwnd,
            Pid = candidate.Pid,
            ThreadId = candidate.ThreadId,
            OwnerHwnd = candidate.OwnerHwnd,
            RootOwnerHwnd = candidate.RootOwnerHwnd,
            Title = candidate.Title,
            ClassName = candidate.ClassName,
            Visible = candidate.Visible,
            ProcessImage = candidate.ProcessImage,
            AutomationId = uia?.AutomationId,
            Name = uia?.Name,
            ControlType = uia?.ControlType,
            Text = candidate.Text.ToArray(),
            Buttons = candidate.Buttons.Select(ToElementSnapshot).ToArray(),
            Children = candidate.Children.Select(ToElementSnapshot).ToArray(),
            Action = action.Name,
            ActionMethod = actionResult.Method,
            ActionTarget = action.TargetText,
            ActionSucceeded = actionResult.Succeeded,
        };
    }

    private static DialogElementSnapshot ToElementSnapshot(WindowElement candidate)
    {
        return new DialogElementSnapshot(candidate.Hwnd, candidate.ClassName, candidate.Text);
    }

    private static bool BelongsToTarget(WindowCandidate candidate, DialogWatchRequest request)
    {
        if (candidate.Pid == request.ExcelProcessId)
        {
            return true;
        }
        if (request.VbeThreadId > 0 && candidate.ThreadId == request.VbeThreadId)
        {
            return true;
        }
        return candidate.OwnerChain.Contains(request.ExcelMainHwnd) ||
               (request.VbeHwnd != 0 && candidate.OwnerChain.Contains(request.VbeHwnd));
    }

    private static bool Matches(DialogKind requested, DialogKind actual)
    {
        return requested == DialogKind.Any || requested == actual;
    }
}

internal static class DialogFingerprint
{
    public static DialogKind? Classify(WindowCandidate candidate, UiaDialogDescription? uia)
    {
        var title = candidate.Title;
        var className = candidate.ClassName;
        var text = string.Join("\n", candidate.Text);
        var buttons = string.Join("\n", candidate.Buttons.Select(button => button.Text));
        var combined = string.Join("\n", title, text, buttons, uia?.Name ?? "");
        var dialogLike = className.Equals("#32770", StringComparison.OrdinalIgnoreCase) ||
                         className.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase) ||
                         className.Equals("NUIDialog", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(uia?.ControlType, "Window", StringComparison.OrdinalIgnoreCase);
        if (!dialogLike)
        {
            return null;
        }

        if (ContainsAny(combined, "run-time error", "runtime error", "実行時エラー") ||
            ContainsAny(buttons, "Debug", "End", "Continue", "デバッグ", "終了", "継続"))
        {
            return DialogKind.Runtime;
        }
        if (ContainsAny(combined, "compile", "syntax error", "expected", "コンパイル", "構文エラー", "必要です"))
        {
            return DialogKind.Compile;
        }
        if (className.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase) ||
            ContainsAny(combined, "Open", "Save As", "Browse", "開く", "名前を付けて保存", "参照"))
        {
            return DialogKind.FileDialog;
        }
        if (candidate.Children.Any(child => child.ClassName.Equals("Edit", StringComparison.OrdinalIgnoreCase)))
        {
            return DialogKind.InputBox;
        }
        if (candidate.Buttons.Count > 0)
        {
            return DialogKind.MsgBox;
        }
        return null;
    }

    private static bool ContainsAny(string value, params string[] tokens)
    {
        return tokens.Any(token => value.Contains(token, StringComparison.OrdinalIgnoreCase));
    }
}

internal static class DialogActionSelector
{
    public static DialogAction Select(DialogKind kind, WindowCandidate candidate, DialogActionPolicy policy)
    {
        if (policy == DialogActionPolicy.ObserveOnly)
        {
            return DialogAction.None;
        }

        if (kind == DialogKind.Runtime && policy == DialogActionPolicy.SuppressVbaError)
        {
            return FindButton(candidate, "runtime_end", "End", "終了") ??
                   FindButton(candidate, "runtime_debug", "Debug", "デバッグ") ??
                   FindButton(candidate, "runtime_close", "OK", "Close", "閉じる") ??
                   new DialogAction("runtime_close", 0, "", true);
        }

        if (kind == DialogKind.Compile && policy == DialogActionPolicy.SuppressVbaError)
        {
            return FindButton(candidate, "compile_close", "OK", "Close", "閉じる") ??
                   DialogAction.None;
        }

        if (policy is DialogActionPolicy.SuppressVbaError or DialogActionPolicy.CancelSupportedNativeUi &&
            kind is DialogKind.MsgBox or DialogKind.InputBox or DialogKind.FileDialog)
        {
            return FindButton(candidate, "native_cancel", "Cancel", "Close", "キャンセル", "閉じる") ??
                   DialogAction.None;
        }

        return DialogAction.None;
    }

    private static DialogAction? FindButton(WindowCandidate candidate, string action, params string[] labels)
    {
        foreach (var button in candidate.Buttons)
        {
            var buttonText = NormalizeButtonText(button.Text);
            if (labels.Any(label => string.Equals(buttonText, NormalizeButtonText(label), StringComparison.OrdinalIgnoreCase)))
            {
                return new DialogAction(action, button.Hwnd, button.Text, false);
            }
        }
        return null;
    }

    private static string NormalizeButtonText(string value)
    {
        var text = value.Trim().Replace("&", "", StringComparison.Ordinal);
        var accelerator = text.LastIndexOf('(');
        if (accelerator > 0 && text.EndsWith(')') && text.Length - accelerator <= 4)
        {
            text = text[..accelerator].TrimEnd();
        }
        return text;
    }
}

internal sealed record DialogAction(string Name, long TargetHwnd, string TargetText, bool AllowWindowClose)
{
    public static DialogAction None { get; } = new("", 0, "", false);
}

internal sealed record DialogActionResult(bool Succeeded, string Method)
{
    public static DialogActionResult None { get; } = new(false, "");
}

internal sealed record WindowElement(long Hwnd, string ClassName, string Text);

internal sealed record WindowCandidate(
    long Hwnd,
    int Pid,
    int ThreadId,
    long OwnerHwnd,
    long RootOwnerHwnd,
    string Title,
    string ClassName,
    bool Visible,
    string ProcessImage,
    IReadOnlyList<long> OwnerChain,
    IReadOnlyList<string> Text,
    IReadOnlyList<WindowElement> Buttons,
    IReadOnlyList<WindowElement> Children);

internal sealed record UiaDialogDescription(string? AutomationId, string? Name, string? ControlType);

internal interface IWindowEnumerator
{
    IReadOnlyList<WindowCandidate> Enumerate(int processId, int vbeThreadId);
    bool ClickButton(long hwnd);
    bool CloseWindow(long hwnd);
}

internal interface IUiaDialogAdapter
{
    UiaDialogDescription? Describe(long hwnd);
    bool TryInvoke(long hwnd);
}

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Win32 window enumeration is Windows-only bridge behavior.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Window process metadata is best-effort diagnostic data.")]
internal sealed class Win32WindowEnumerator : IWindowEnumerator
{
    public IReadOnlyList<WindowCandidate> Enumerate(int processId, int vbeThreadId)
    {
        var handles = new HashSet<long>();
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            handles.Add(hwnd.ToInt64());
            return true;
        }, IntPtr.Zero);
        if (vbeThreadId > 0)
        {
            NativeMethods.EnumThreadWindows((uint)vbeThreadId, (hwnd, _) =>
            {
                handles.Add(hwnd.ToInt64());
                return true;
            }, IntPtr.Zero);
        }

        var result = new List<WindowCandidate>();
        foreach (var handle in handles)
        {
            var hwnd = new IntPtr(handle);
            _ = NativeMethods.GetWindowThreadProcessId(hwnd, out var pid);
            var threadId = unchecked((int)NativeMethods.GetWindowThreadProcessId(hwnd, out _));
            var owner = NativeMethods.GetWindow(hwnd, NativeMethods.GwOwner).ToInt64();
            var ownerChain = GetOwnerChain(hwnd);
            var rootOwner = ownerChain.Count == 0 ? 0 : ownerChain[^1];
            var children = GetChildren(hwnd);
            var buttons = children.Where(child => child.ClassName.Equals("Button", StringComparison.OrdinalIgnoreCase)).ToArray();
            var text = children
                .Where(child => child.ClassName.Equals("Static", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(child.Text))
                .Select(child => child.Text)
                .ToArray();
            result.Add(new WindowCandidate(
                handle,
                pid,
                threadId,
                owner,
                rootOwner,
                NativeMethods.GetWindowText(hwnd),
                NativeMethods.GetClassName(hwnd),
                NativeMethods.IsWindowVisible(hwnd),
                GetProcessImage(pid),
                ownerChain,
                text,
                buttons,
                children));
        }
        return result;
    }

    public bool ClickButton(long hwnd)
    {
        return NativeMethods.SendMessage(new IntPtr(hwnd), NativeMethods.BmClick, IntPtr.Zero, IntPtr.Zero) != IntPtr.Zero ||
               NativeMethods.IsWindow(new IntPtr(hwnd));
    }

    public bool CloseWindow(long hwnd)
    {
        return NativeMethods.PostMessage(new IntPtr(hwnd), NativeMethods.WmClose, IntPtr.Zero, IntPtr.Zero);
    }

    private static List<long> GetOwnerChain(IntPtr hwnd)
    {
        var owners = new List<long>();
        var current = NativeMethods.GetWindow(hwnd, NativeMethods.GwOwner);
        while (current != IntPtr.Zero && owners.Count < 16)
        {
            owners.Add(current.ToInt64());
            current = NativeMethods.GetWindow(current, NativeMethods.GwOwner);
        }
        return owners;
    }

    private static List<WindowElement> GetChildren(IntPtr hwnd)
    {
        var children = new List<WindowElement>();
        NativeMethods.EnumChildWindows(hwnd, (child, _) =>
        {
            children.Add(new WindowElement(child.ToInt64(), NativeMethods.GetClassName(child), NativeMethods.GetWindowText(child)));
            return true;
        }, IntPtr.Zero);
        return children;
    }

    private static string GetProcessImage(int pid)
    {
        try
        {
            return Process.GetProcessById(pid).MainModule?.FileName ?? "";
        }
        catch
        {
            return "";
        }
    }
}

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "UI Automation is Windows-only bridge behavior.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "UI Automation metadata and action attempts are best-effort fallbacks.")]
internal sealed class ComUiaDialogAdapter : IUiaDialogAdapter
{
    private readonly object? _automation;

    public ComUiaDialogAdapter()
    {
        try
        {
            var type = Type.GetTypeFromProgID("UIAutomationClient.CUIAutomation") ??
                       Type.GetTypeFromProgID("CUIAutomation");
            _automation = type is null ? null : Activator.CreateInstance(type);
        }
        catch
        {
            _automation = null;
        }
    }

    public UiaDialogDescription? Describe(long hwnd)
    {
        if (_automation is null || hwnd == 0)
        {
            return null;
        }
        try
        {
            dynamic automation = _automation;
            dynamic element = automation.ElementFromHandle(new IntPtr(hwnd));
            return new UiaDialogDescription(
                Convert.ToString(element.CurrentAutomationId),
                Convert.ToString(element.CurrentName),
                Convert.ToString(element.CurrentControlType));
        }
        catch
        {
            return null;
        }
    }

    public bool TryInvoke(long hwnd)
    {
        if (_automation is null || hwnd == 0)
        {
            return false;
        }
        try
        {
            dynamic automation = _automation;
            dynamic element = automation.ElementFromHandle(new IntPtr(hwnd));
            const int uiAInvokePatternId = 10000;
            dynamic pattern = element.GetCurrentPattern(uiAInvokePatternId);
            pattern.Invoke();
            return true;
        }
        catch
        {
            return false;
        }
    }
}

[SuppressMessage("Performance", "CA1838:Avoid StringBuilder parameters for P/Invokes", Justification = "The bounded diagnostic text helpers favor simple and compatible Win32 signatures.")]
internal static class NativeMethods
{
    public const uint GwOwner = 4;
    public const uint BmClick = 0x00F5;
    public const uint WmClose = 0x0010;

    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumThreadWindows(uint threadId, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumChildWindows(IntPtr hwnd, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int processId);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hwnd, uint command);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder className, int maxCount);

    public static string GetWindowText(IntPtr hwnd)
    {
        var text = new StringBuilder(1024);
        _ = GetWindowText(hwnd, text, text.Capacity);
        return text.ToString();
    }

    public static string GetClassName(IntPtr hwnd)
    {
        var text = new StringBuilder(256);
        _ = GetClassName(hwnd, text, text.Capacity);
        return text.ToString();
    }
}
