namespace Xlflow.ExcelBridge.Services;

public sealed record InspectCommandArguments(
    string Target,
    string Sheet,
    string Address,
    string WorkbookPath,
    string MetadataPath,
    bool UseSession,
    bool IncludeStyle,
    int MaxRows,
    int MaxCols);
