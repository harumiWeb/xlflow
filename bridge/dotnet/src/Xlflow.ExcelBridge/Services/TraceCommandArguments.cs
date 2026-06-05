namespace Xlflow.ExcelBridge.Services;

public sealed record TraceCommandArguments(
    string Action,
    string WorkbookPath,
    string ModulesDir,
    bool Visible,
    bool Force,
    string TraceDir,
    bool UseSession,
    string MetadataPath);
