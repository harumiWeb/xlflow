namespace Xlflow.ExcelBridge.Services;

public sealed record InspectFormCommandArguments(
    string WorkbookPath,
    string FormName,
    string Basis,
    string Initializer,
    bool StrictDesigner,
    bool Visible,
    bool UseSession,
    string MetadataPath);
