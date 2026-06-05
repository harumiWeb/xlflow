namespace Xlflow.ExcelBridge.Services;

public sealed record ExportImageCommandArguments(
    string WorkbookPath,
    string Sheet,
    string RangeAddress,
    string OutputPath,
    bool OutputIsDefault,
    string ImageFormat,
    bool Overwrite,
    bool Visible,
    bool UseSession,
    string MetadataPath);
