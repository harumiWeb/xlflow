namespace Xlflow.ExcelBridge.Services;

public sealed record RunCommandArguments(
    string WorkbookPath,
    string MacroName,
    string MacroArgsJSON,
    bool Visible,
    bool DisplayAlerts,
    bool SaveWorkbook,
    bool Diagnostic,
    bool SuppressModalErrors,
    bool UseSession,
    string MetadataPath,
    string RuntimeMode,
    string RuntimeSource,
    string SaveAsPath,
    int TimeoutSeconds);
