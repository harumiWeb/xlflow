namespace Xlflow.ExcelBridge.Services;

public sealed record MacrosCommandArguments(
    string WorkbookPath,
    string MetadataPath,
    bool UseSession,
    bool Visible,
    string Entry,
    bool RunnableOnly);
