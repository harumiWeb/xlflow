using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text.Json.Serialization;

namespace Xlflow.ExcelBridge.Services;

internal sealed record ErrorLocation(
    string Confidence,
    string Method,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? SourcePath,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Component,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? ComponentType,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Procedure,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? Line,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? Column,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? EndLine,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? EndColumn,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? Text);

internal sealed record VbeSelectionCaptureAttempt(
    string Timing,
    bool Success,
    string? Error = null);

internal sealed record VbeSelectionCapture(
    ErrorLocation? Location,
    IReadOnlyList<VbeSelectionCaptureAttempt> Attempts)
{
    public static VbeSelectionCapture Empty { get; } = new(null, []);

    public bool HasMeaningfulLocation =>
        Location is not null &&
        (!string.IsNullOrWhiteSpace(Location.Component) ||
         Location.Line is > 0 ||
         !string.IsNullOrWhiteSpace(Location.SourcePath));

    public bool HasReliableLocation =>
        Location is not null &&
        VbeSelectionScorer.Score(Location) >= VbeSelectionScorer.ReliableThreshold;
}

internal interface IVbeSelectionLocator
{
    VbeSelectionCapture Capture(string timing);
    void ResetBreakMode();
}

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "VBE selection capture is best-effort diagnostics and must not fail the bridge command.")]
internal sealed class VbeSelectionLocator(
    int excelProcessId,
    long excelHwnd,
    VbeSourceMappingOptions sourceOptions,
    TimeSpan? timeout = null) : IVbeSelectionLocator
{
    private const int ResetControlId = 228;
    private readonly TimeSpan _timeout = timeout ?? TimeSpan.FromMilliseconds(750);

    public VbeSelectionCapture Capture(string timing)
    {
        try
        {
            var task = Task.Run(() => CaptureCore(timing));
            if (task.Wait(_timeout))
            {
                return task.Result;
            }
            return new VbeSelectionCapture(null, [new VbeSelectionCaptureAttempt(timing, false, "VBE selection capture timed out.")]);
        }
        catch (Exception ex)
        {
            return new VbeSelectionCapture(null, [new VbeSelectionCaptureAttempt(timing, false, ExcelBridgeSupport.FormatExceptionDetail(ex))]);
        }
    }

    public void ResetBreakMode()
    {
        object? excel = null;
        object? vbe = null;
        object? control = null;
        try
        {
            excel = excelHwnd != 0
                ? ExcelBridgeSupport.TryGetExcelByHwnd(excelHwnd)
                : null;
            excel ??= ExcelBridgeSupport.TryGetExcelByProcessId(excelProcessId);
            if (excel is null)
            {
                return;
            }

            vbe = ExcelBridgeSupport.Get(excel, "VBE");
            if (vbe is null)
            {
                return;
            }

            control = FindResetControl(vbe);
            if (control is not null)
            {
                ExcelBridgeSupport.InvokeMethod(control, "Execute");
            }
        }
        catch
        {
            // Break-mode reset is best-effort; diagnostics must remain non-fatal.
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(control);
            ExcelBridgeSupport.ReleaseComObject(vbe);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private VbeSelectionCapture CaptureCore(string timing)
    {
        object? excel = null;
        object? vbe = null;
        try
        {
            excel = excelHwnd != 0
                ? ExcelBridgeSupport.TryGetExcelByHwnd(excelHwnd)
                : null;
            excel ??= ExcelBridgeSupport.TryGetExcelByProcessId(excelProcessId);
            if (excel is null)
            {
                return Failed(timing, "Excel instance was not available.");
            }

            vbe = ExcelBridgeSupport.Get(excel, "VBE");
            if (vbe is null)
            {
                return Failed(timing, "VBE object was not available.");
            }

            var locations = new List<ErrorLocation>();
            object? activePane = null;
            try
            {
                activePane = ExcelBridgeSupport.Get(vbe, "ActiveCodePane");
                if (activePane is not null)
                {
                    if (CapturePane(activePane) is { } activeLocation)
                    {
                        locations.Add(activeLocation);
                    }
                }
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(activePane);
            }

            foreach (var codePane in EnumerateCodePanes(vbe))
            {
                try
                {
                    if (CapturePane(codePane) is { } location)
                    {
                        locations.Add(location);
                    }
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(codePane);
                }
            }

            var best = locations
                .OrderByDescending(VbeSelectionScorer.Score)
                .FirstOrDefault();
            if (best is null)
            {
                return Failed(timing, "VBE ActiveCodePane was not available.");
            }

            return new VbeSelectionCapture(best, [new VbeSelectionCaptureAttempt(timing, true)]);
        }
        catch (Exception ex)
        {
            return Failed(timing, ExcelBridgeSupport.FormatExceptionDetail(ex));
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(vbe);
            ExcelBridgeSupport.ReleaseComObject(excel);
        }
    }

    private ErrorLocation? CapturePane(object pane)
    {
        object? codeModule = null;
        object? component = null;
        try
        {
            codeModule = ExcelBridgeSupport.Get(pane, "CodeModule");
            if (codeModule is null)
            {
                return null;
            }

            component = ExcelBridgeSupport.Get(codeModule, "Parent");
            var componentName = TryGetString(component, "Name");
            var componentType = TryGetInt(component, "Type");
            var componentTypeName = ComponentTypeName(componentType);
            var sourcePathInfo = VbeSourcePathMapper.ResolveSourcePathInfo(component, sourceOptions);
            var sourcePath = sourcePathInfo?.JsonPath;
            var selection = TryGetSelection(pane);
            if (selection is not null && !string.IsNullOrWhiteSpace(componentName))
            {
                var text = TryGetLine(codeModule, selection.Value.StartLine);
                var procedure = TryGetProcedure(codeModule, selection.Value.StartLine);
                var location = new ErrorLocation(
                    "high",
                    "vbe.selection",
                    sourcePath,
                    componentName,
                    componentTypeName,
                    procedure,
                    selection.Value.StartLine,
                    selection.Value.StartColumn,
                    selection.Value.EndLine,
                    selection.Value.EndColumn,
                    text);
                return MapLocationToSourceFile(RefineLowValueSelection(codeModule, location), sourcePathInfo?.FullPath);
            }

            if (!string.IsNullOrWhiteSpace(componentName))
            {
                return new ErrorLocation(
                    "medium",
                    "vbe.active_code_pane",
                    sourcePath,
                    componentName,
                    componentTypeName,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            }

            return null;
        }
        catch
        {
            return null;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(component);
            ExcelBridgeSupport.ReleaseComObject(codeModule);
        }
    }

    private static List<object> EnumerateCodePanes(object vbe)
    {
        var result = new List<object>();
        object? codePanes = null;
        try
        {
            codePanes = ExcelBridgeSupport.Get(vbe, "CodePanes");
            if (codePanes is null)
            {
                return result;
            }
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(codePanes, "Count"));
            for (var index = 1; index <= count; index++)
            {
                var pane = ExcelBridgeSupport.Get(codePanes, "Item", index);
                if (pane is not null)
                {
                    result.Add(pane);
                }
            }
        }
        catch
        {
            return result;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(codePanes);
        }
        return result;
    }

    private static VbeSelectionCapture Failed(string timing, string error)
    {
        return new VbeSelectionCapture(null, [new VbeSelectionCaptureAttempt(timing, false, error)]);
    }

    private static (int StartLine, int StartColumn, int EndLine, int EndColumn)? TryGetSelection(object pane)
    {
        try
        {
            dynamic activePane = pane;
            var startLine = 0;
            var startColumn = 0;
            var endLine = 0;
            var endColumn = 0;
            activePane.GetSelection(ref startLine, ref startColumn, ref endLine, ref endColumn);
            if (startLine <= 0 || startColumn <= 0)
            {
                return null;
            }
            return (startLine, startColumn, endLine > 0 ? endLine : startLine, endColumn > 0 ? endColumn : startColumn);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetLine(object codeModule, int line)
    {
        if (line <= 0)
        {
            return null;
        }
        try
        {
            return ExcelBridgeSupport.GetString(codeModule, "Lines", line, 1);
        }
        catch
        {
            return null;
        }
    }

    private static ErrorLocation RefineLowValueSelection(object codeModule, ErrorLocation location)
    {
        if (!VbeSelectionScorer.IsLowValueText(location.Text))
        {
            return location;
        }

        if (FindLikelyCompileErrorLine(codeModule) is not { } candidate)
        {
            return location;
        }

        var text = TryGetLine(codeModule, candidate);
        if (string.IsNullOrWhiteSpace(text))
        {
            return location;
        }

        var column = FirstNonWhitespaceColumn(text);
        return location with
        {
            Procedure = TryGetProcedure(codeModule, candidate),
            Line = candidate,
            Column = column,
            EndLine = candidate,
            EndColumn = Math.Max(column, text.Length + 1),
            Text = text,
        };
    }

    private static ErrorLocation MapLocationToSourceFile(ErrorLocation location, string? sourceFullPath)
    {
        if (location.Line is not > 0 || string.IsNullOrWhiteSpace(sourceFullPath) || !File.Exists(sourceFullPath))
        {
            return location;
        }

        try
        {
            var mappedLine = SourceLineMapper.MapVbeLineToSourceLine(
                File.ReadAllText(sourceFullPath),
                location.Line.Value,
                location.Text);
            if (mappedLine is null)
            {
                return location with
                {
                    Line = null,
                    Column = null,
                    EndLine = null,
                    EndColumn = null,
                };
            }

            var lineDelta = mappedLine.Value - location.Line.Value;
            return location with
            {
                Line = mappedLine,
                Column = ReliableSourceColumn(location),
                EndLine = location.EndLine is > 0 ? location.EndLine + lineDelta : mappedLine,
                EndColumn = ReliableSourceColumn(location) is not null ? location.EndColumn : null,
            };
        }
        catch
        {
            return location with
            {
                Line = null,
                Column = null,
                EndLine = null,
                EndColumn = null,
            };
        }

    }

    private static int? ReliableSourceColumn(ErrorLocation location)
    {
        if (location.Column is not > 0)
        {
            return null;
        }

        if (location.Column == 1 && !string.IsNullOrEmpty(location.Text) && char.IsWhiteSpace(location.Text[0]))
        {
            return null;
        }

        return location.Column;
    }

    private static int? FindLikelyCompileErrorLine(object codeModule)
    {
        try
        {
            var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(codeModule, "CountOfLines"));
            for (var line = 1; line <= count; line++)
            {
                var text = TryGetLine(codeModule, line);
                if (VbeSelectionScorer.IsLikelyCompileErrorLine(text))
                {
                    return line;
                }
            }
        }
        catch
        {
            return null;
        }
        return null;
    }

    private static int FirstNonWhitespaceColumn(string text)
    {
        for (var index = 0; index < text.Length; index++)
        {
            if (!char.IsWhiteSpace(text[index]))
            {
                return index + 1;
            }
        }
        return 1;
    }

    private static string? TryGetProcedure(object codeModule, int line)
    {
        if (line <= 0)
        {
            return null;
        }
        try
        {
            dynamic module = codeModule;
            var kind = 0;
            var procedure = Convert.ToString(module.ProcOfLine(line, ref kind), CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(procedure) ? null : procedure;
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetString(object? comObject, string memberName)
    {
        if (comObject is null)
        {
            return null;
        }
        try
        {
            return ExcelBridgeSupport.GetString(comObject, memberName);
        }
        catch
        {
            return null;
        }
    }

    private static int TryGetInt(object? comObject, string memberName)
    {
        if (comObject is null)
        {
            return 0;
        }
        try
        {
            return ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(comObject, memberName));
        }
        catch
        {
            return 0;
        }
    }

    private static string? ComponentTypeName(int type)
    {
        return type switch
        {
            1 => "module",
            2 => "class",
            3 => "form",
            100 => "document",
            _ => null,
        };
    }

    private static object? FindResetControl(object vbe)
    {
        object? commandBars = null;
        try
        {
            commandBars = ExcelBridgeSupport.Get(vbe, "CommandBars");
            if (commandBars is null)
            {
                return null;
            }

            var byId = FindBuiltInControl(commandBars, ResetControlId);
            if (byId is not null)
            {
                return byId;
            }

            foreach (var barName in new[] { "Run", "実行" })
            {
                object? bar = null;
                object? controls = null;
                try
                {
                    bar = ExcelBridgeSupport.Get(commandBars, "Item", barName);
                    if (bar is null)
                    {
                        continue;
                    }
                    controls = ExcelBridgeSupport.Get(bar, "Controls");
                    if (controls is null)
                    {
                        continue;
                    }

                    var byCaption = FindResetControlByCaption(controls, "Reset") ??
                                    FindResetControlByCaption(controls, "リセット");
                    if (byCaption is not null)
                    {
                        return byCaption;
                    }

                    var count = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(controls, "Count"));
                    for (var index = 1; index <= count; index++)
                    {
                        object? candidate = null;
                        try
                        {
                            candidate = ExcelBridgeSupport.Get(controls, "Item", index);
                            var caption = (ExcelBridgeSupport.GetString(candidate!, "Caption") ?? "").Replace("&", "", StringComparison.Ordinal);
                            if (caption.Contains("Reset", StringComparison.OrdinalIgnoreCase) ||
                                caption.Contains("リセット", StringComparison.Ordinal))
                            {
                                var found = candidate;
                                candidate = null;
                                return found;
                            }
                        }
                        catch
                        {
                            // Try the next control.
                        }
                        finally
                        {
                            ExcelBridgeSupport.ReleaseComObject(candidate);
                        }
                    }
                }
                catch
                {
                    // Try the next localized command bar.
                }
                finally
                {
                    ExcelBridgeSupport.ReleaseComObject(controls);
                    ExcelBridgeSupport.ReleaseComObject(bar);
                }
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            ExcelBridgeSupport.ReleaseComObject(commandBars);
        }
        return null;
    }

    private static object? FindBuiltInControl(object commandBars, int controlId)
    {
        try
        {
            return ExcelBridgeSupport.InvokeMethod(commandBars, "FindControl", Type.Missing, controlId);
        }
        catch
        {
            return null;
        }
    }

    private static object? FindResetControlByCaption(object controls, string caption)
    {
        try
        {
            return ExcelBridgeSupport.Get(controls, "Item", caption);
        }
        catch
        {
            return null;
        }
    }
}

internal static class VbeSelectionScorer
{
    public const int ReliableThreshold = 90;

    public static int Score(ErrorLocation location)
    {
        var score = string.Equals(location.Method, "vbe.selection", StringComparison.OrdinalIgnoreCase) ? 100 : 50;
        if (!string.IsNullOrWhiteSpace(location.SourcePath))
        {
            score += 20;
        }
        if (!string.IsNullOrWhiteSpace(location.Procedure))
        {
            score += 10;
        }
        if (location.Line is > 0)
        {
            score += 10;
        }
        if (!string.IsNullOrWhiteSpace(location.Text))
        {
            score += 10;
        }
        if (IsLowValueText(location.Text))
        {
            score -= 80;
        }
        if (IsTemporaryHarness(location.Component))
        {
            score -= 1000;
        }
        return score;
    }

    private static bool IsTemporaryHarness(string? component)
    {
        return !string.IsNullOrWhiteSpace(component) &&
               component.StartsWith("XlflowRun_", StringComparison.OrdinalIgnoreCase);
    }

    internal static bool IsLowValueText(string? text)
    {
        var trimmed = text?.Trim() ?? "";
        return string.IsNullOrWhiteSpace(trimmed) ||
               trimmed.StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("Option ", StringComparison.OrdinalIgnoreCase);
    }

    internal static bool IsLikelyCompileErrorLine(string? text)
    {
        var trimmed = text?.Trim() ?? "";
        if (trimmed.Length == 0 ||
            trimmed.StartsWith('\'') ||
            trimmed.StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Option ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Public Sub ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Private Sub ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("End ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Dim ", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return trimmed.EndsWith('=') ||
               trimmed.EndsWith('.') ||
               trimmed.EndsWith(',') ||
               trimmed.EndsWith('+') ||
               trimmed.EndsWith('-') ||
               trimmed.EndsWith('*') ||
               trimmed.EndsWith('/');
    }
}

internal static class SourceLineMapper
{
    public static int? MapVbeLineToSourceLine(string sourceText, int vbeLine, string? vbeText)
    {
        if (vbeLine <= 0)
        {
            return null;
        }

        var visibleLines = SplitLines(sourceText)
            .Select((text, index) => new SourceVisibleLine(index + 1, text))
            .Where(line => !IsHiddenExportMetadata(line.Text))
            .ToArray();
        if (vbeLine > visibleLines.Length)
        {
            return null;
        }

        var mapped = visibleLines[vbeLine - 1];
        if (IsSameCodeLine(mapped.Text, vbeText))
        {
            return mapped.SourceLine;
        }

        if (!string.IsNullOrWhiteSpace(vbeText))
        {
            var normalizedVbeText = NormalizeCodeLine(vbeText);
            var matchingLines = visibleLines
                .Where(line => string.Equals(NormalizeCodeLine(line.Text), normalizedVbeText, StringComparison.Ordinal))
                .Select(line => line.SourceLine)
                .ToArray();
            if (matchingLines.Length == 1)
            {
                return matchingLines[0];
            }
        }

        return null;
    }

    private static bool IsHiddenExportMetadata(string text)
    {
        var trimmed = text.Trim();
        return trimmed.StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "VERSION 1.0 CLASS", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "VERSION 5.00", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "BEGIN", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "END", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("MultiUse ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSameCodeLine(string sourceText, string? vbeText)
    {
        if (string.IsNullOrWhiteSpace(vbeText))
        {
            return true;
        }

        return string.Equals(NormalizeCodeLine(sourceText), NormalizeCodeLine(vbeText), StringComparison.Ordinal);
    }

    private static string NormalizeCodeLine(string? text)
    {
        return (text ?? "").Trim();
    }

    private static string[] SplitLines(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return [];
        }
        return text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
    }

    private sealed record SourceVisibleLine(int SourceLine, string Text);
}

internal sealed record VbeSourceMappingOptions(
    string ModulesDir,
    string ClassesDir,
    string FormsDir,
    string WorkbookDir,
    string CodeSource,
    bool Folders,
    string FolderAnnotation,
    bool DefaultComponentFolders);

internal sealed record VbeSourcePathInfo(string JsonPath, string FullPath);

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Source path mapping is best-effort diagnostic enrichment.")]
internal static class VbeSourcePathMapper
{
    public static string? ResolveSourcePath(object? component, VbeSourceMappingOptions options)
    {
        return ResolveSourcePathInfo(component, options)?.JsonPath;
    }

    public static VbeSourcePathInfo? ResolveSourcePathInfo(object? component, VbeSourceMappingOptions options)
    {
        if (component is null)
        {
            return null;
        }

        try
        {
            var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));
            var name = ExcelBridgeSupport.GetString(component, "Name") ?? "";
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            string? path = null;
            if (type == 3 && VbaSourceHelper.IsSidecarMode(options.CodeSource))
            {
                var codePath = VbaSourceHelper.GetUserFormCodePath(options.FormsDir, name);
                if (!string.IsNullOrWhiteSpace(codePath) && File.Exists(codePath))
                {
                    path = codePath;
                }
            }

            path ??= VbaSourceHelper.GetComponentPath(
                component,
                options.ModulesDir,
                options.ClassesDir,
                options.FormsDir,
                options.WorkbookDir,
                options.Folders,
                string.IsNullOrWhiteSpace(options.FolderAnnotation) ? "update" : options.FolderAnnotation,
                options.DefaultComponentFolders);

            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }
            return new VbeSourcePathInfo(ToProjectRelativePath(path, options) ?? path, path);
        }
        catch
        {
            return null;
        }
    }

    internal static string? ToProjectRelativePath(string? path, VbeSourceMappingOptions options)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        var root = ResolveProjectRoot(options);
        if (string.IsNullOrWhiteSpace(root))
        {
            return path;
        }

        try
        {
            var fullRoot = Path.GetFullPath(root);
            var fullPath = Path.GetFullPath(path);
            var relative = Path.GetRelativePath(fullRoot, fullPath);
            if (Path.IsPathRooted(relative) ||
                relative.Equals("..", StringComparison.Ordinal) ||
                relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) ||
                relative.StartsWith($"..{Path.AltDirectorySeparatorChar}", StringComparison.Ordinal))
            {
                return path;
            }
            return relative.Replace('\\', '/');
        }
        catch
        {
            return path;
        }
    }

    private static string? ResolveProjectRoot(VbeSourceMappingOptions options)
    {
        var dirs = new[]
            {
                options.ModulesDir,
                options.ClassesDir,
                options.FormsDir,
                options.WorkbookDir,
            }
            .Where(dir => !string.IsNullOrWhiteSpace(dir))
            .Select(Path.GetFullPath)
            .ToArray();
        if (dirs.Length == 0)
        {
            return null;
        }

        var common = dirs[0];
        foreach (var dir in dirs.Skip(1))
        {
            common = CommonDirectory(common, dir);
        }

        var commonInfo = new DirectoryInfo(common);
        if (string.Equals(commonInfo.Name, "src", StringComparison.OrdinalIgnoreCase) &&
            commonInfo.Parent is not null)
        {
            return commonInfo.Parent.FullName;
        }

        return commonInfo.FullName;
    }

    private static string CommonDirectory(string left, string right)
    {
        var leftFull = Path.GetFullPath(left);
        var rightFull = Path.GetFullPath(right);
        var leftRoot = Path.GetPathRoot(leftFull) ?? "";
        var rightRoot = Path.GetPathRoot(rightFull) ?? "";
        if (!string.Equals(leftRoot, rightRoot, StringComparison.OrdinalIgnoreCase))
        {
            return leftFull;
        }

        var leftParts = SplitPath(leftFull, leftRoot);
        var rightParts = SplitPath(rightFull, rightRoot);
        var count = Math.Min(leftParts.Length, rightParts.Length);
        var common = new List<string>();
        for (var i = 0; i < count; i++)
        {
            if (!string.Equals(leftParts[i], rightParts[i], StringComparison.OrdinalIgnoreCase))
            {
                break;
            }
            common.Add(leftParts[i]);
        }
        return common.Count == 0 ? leftRoot : Path.Combine([leftRoot, .. common]);
    }

    private static string[] SplitPath(string path, string root)
    {
        var relative = path.StartsWith(root, StringComparison.OrdinalIgnoreCase)
            ? path[root.Length..]
            : path;
        return relative
            .Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries);
    }
}
