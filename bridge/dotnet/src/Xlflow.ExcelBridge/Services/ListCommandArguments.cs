namespace Xlflow.ExcelBridge.Services;

public sealed record ListCommandArguments(
    string Action,
    string WorkbookPath,
    string FormsDir,
    string ProjectRoot,
    bool Visible,
    bool UseSession,
    string MetadataPath);
