namespace Xlflow.ExcelBridge.Services;

public sealed record UICommandArguments(
    string Action,
    string WorkbookPath,
    bool Visible,
    string Sheet,
    string Cell,
    string Text,
    string Macro,
    string Id,
    int Width,
    int Height,
    bool CreateSheet,
    bool VerifyMacro,
    bool UseSession,
    string MetadataPath);
