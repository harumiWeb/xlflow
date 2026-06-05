namespace Xlflow.ExcelBridge.Services;

public sealed record FormExportImageCommandArguments(
    string WorkbookPath,
    string FormName,
    string OutputPath,
    string Initializer,
    bool Overwrite,
    bool Visible,
    bool UseSession,
    string MetadataPath);
