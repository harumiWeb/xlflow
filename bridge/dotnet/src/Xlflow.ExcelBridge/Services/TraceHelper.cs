using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Bridge helpers intentionally normalize COM failures into best-effort results.")]
internal static class TraceHelper
{
    public static string BuildTraceModuleCode()
    {
        var builder = new StringBuilder();
        builder.AppendLine("Option Explicit");
        builder.AppendLine();
        builder.AppendLine("Private mTraceFile As String");
        builder.AppendLine();
        builder.AppendLine("Public Sub XlflowSetTraceFile(ByVal path As String)");
        builder.AppendLine("  mTraceFile = path");
        builder.AppendLine("End Sub");
        builder.AppendLine();
        builder.AppendLine("Public Sub XlflowLog(ByVal message As String)");
        builder.AppendLine("  If Len(mTraceFile) = 0 Then");
        builder.AppendLine("    Err.Raise vbObjectError + 900, \"XlflowTrace.XlflowLog\", \"trace file is not configured. Run the macro with xlflow run --trace.\"");
        builder.AppendLine("  End If");
        builder.AppendLine("  Dim f As Integer");
        builder.AppendLine("  Dim opened As Boolean");
        builder.AppendLine("  On Error GoTo Handler");
        builder.AppendLine("  f = FreeFile");
        builder.AppendLine("  Open mTraceFile For Append As #f");
        builder.AppendLine("  opened = True");
        builder.AppendLine("  Print #f, Format$(Now, \"yyyy-mm-dd hh:nn:ss\") & vbTab & message");
        builder.AppendLine("  Close #f");
        builder.AppendLine("  Exit Sub");
        builder.AppendLine("Handler:");
        builder.AppendLine("  Dim errNumber As Long");
        builder.AppendLine("  Dim errSource As String");
        builder.AppendLine("  Dim errDescription As String");
        builder.AppendLine("  errNumber = Err.Number");
        builder.AppendLine("  errSource = Err.Source");
        builder.AppendLine("  errDescription = Err.Description");
        builder.AppendLine("  On Error Resume Next");
        builder.AppendLine("  If opened Then Close #f");
        builder.AppendLine("  On Error GoTo 0");
        builder.AppendLine("  Err.Raise errNumber, errSource, errDescription");
        builder.AppendLine("End Sub");
        return builder.ToString();
    }

    public static string GetTraceModuleSourceText()
    {
        return "Attribute VB_Name = \"XlflowTrace\"" + Environment.NewLine + BuildTraceModuleCode();
    }

    public static bool HasTraceModule(object vbProject)
    {
        try
        {
            var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is null)
            {
                return false;
            }
            try
            {
                var existing = ExcelBridgeSupport.Get(components, "Item", "XlflowTrace");
                if (existing is not null)
                {
                    ExcelBridgeSupport.ReleaseComObject(existing);
                    ExcelBridgeSupport.ReleaseComObject(components);
                    return true;
                }
            }
            catch
            {
                // not found
            }
            ExcelBridgeSupport.ReleaseComObject(components);
        }
        catch
        {
            // best-effort
        }
        return false;
    }

    public static bool RemoveTraceModule(object vbProject)
    {
        try
        {
            var components = ExcelBridgeSupport.Get(vbProject, "VBComponents");
            if (components is null)
            {
                return false;
            }
            object? existing = null;
            try
            {
                existing = ExcelBridgeSupport.Get(components, "Item", "XlflowTrace");
            }
            catch
            {
                // not found
            }
            if (existing is not null)
            {
                ExcelBridgeSupport.InvokeMethod(components, "Remove", existing);
                ExcelBridgeSupport.ReleaseComObject(existing);
                ExcelBridgeSupport.ReleaseComObject(components);
                return true;
            }
            ExcelBridgeSupport.ReleaseComObject(components);
        }
        catch
        {
            // best-effort
        }
        return false;
    }

    public static void InstallTraceModule(object vbProject)
    {
        RemoveTraceModule(vbProject);
        var components = ExcelBridgeSupport.Get(vbProject, "VBComponents")
            ?? throw new InvalidOperationException("VBComponents is unavailable.");
        var component = ExcelBridgeSupport.InvokeMethod(components, "Add", 1)
            ?? throw new InvalidOperationException("Could not add trace module.");
        SetProperty(component, "Name", "XlflowTrace");
        var codeModule = ExcelBridgeSupport.Get(component, "CodeModule")
            ?? throw new InvalidOperationException("CodeModule is unavailable.");
        ExcelBridgeSupport.InvokeMethod(codeModule, "AddFromString", BuildTraceModuleCode());
        ExcelBridgeSupport.ReleaseComObject(codeModule);
        ExcelBridgeSupport.ReleaseComObject(component);
        ExcelBridgeSupport.ReleaseComObject(components);
    }

    public static string? WriteTraceModuleSource(string modulesDir)
    {
        if (string.IsNullOrWhiteSpace(modulesDir))
        {
            return null;
        }
        Directory.CreateDirectory(modulesDir);
        var path = Path.Combine(modulesDir, "XlflowTrace.bas");
        File.WriteAllText(path, GetTraceModuleSourceText(), new UTF8Encoding(false));
        return path;
    }

    public static bool TraceModuleSourceMatches(string modulesDir)
    {
        if (string.IsNullOrWhiteSpace(modulesDir))
        {
            return false;
        }
        var path = Path.Combine(modulesDir, "XlflowTrace.bas");
        if (!File.Exists(path))
        {
            return false;
        }
        var existing = File.ReadAllText(path, new UTF8Encoding(false)).Trim();
        var expected = GetTraceModuleSourceText().Trim();
        return existing == expected;
    }

    public static List<TraceEvent> ReadTraceEvents(string path)
    {
        var events = new List<TraceEvent>();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return events;
        }
        var lines = File.ReadAllLines(path);
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }
            events.Add(ParseTraceEvent(line));
        }
        return events;
    }

    private static TraceEvent ParseTraceEvent(string line)
    {
        var tab = line.IndexOf('\t');
        if (tab >= 0)
        {
            var timestamp = line[..tab];
            var message = tab + 1 < line.Length ? line[(tab + 1)..] : "";
            return new TraceEvent(timestamp, message, line);
        }
        return new TraceEvent("", line, line);
    }

    private static void SetProperty(object comObject, string name, object value)
    {
        comObject.GetType().InvokeMember(
            name,
            System.Reflection.BindingFlags.SetProperty,
            null,
            comObject,
            [value],
            System.Globalization.CultureInfo.InvariantCulture);
    }

    public sealed record TraceEvent(string Timestamp, string Message, string Raw);
}
