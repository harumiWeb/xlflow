namespace Xlflow.ExcelBridge.Services;

public sealed record EditCommandArguments(
    string Action,
    string WorkbookPath,
    bool Visible,
    string Sheet,
    string Cell,
    string RangeAddress,
    string Rows,
    string Columns,
    string Value,
    string Formula,
    string Fill,
    string Clear,
    string Height,
    string Width,
    string Events,
    bool UseSession,
    string MetadataPath);
