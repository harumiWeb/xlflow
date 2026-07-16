namespace Xlflow.ExcelBridge.Services;

public sealed record SessionCommandArguments(
    string Action,
    string WorkbookPath,
    string MetadataPath,
    bool Visible,
    bool UseSession,
    bool Discard);
