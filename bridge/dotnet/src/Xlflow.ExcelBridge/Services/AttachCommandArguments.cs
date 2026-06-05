namespace Xlflow.ExcelBridge.Services;

public sealed record AttachCommandArguments(
    string WorkbookPath,
    bool Active,
    string MetadataPath);
