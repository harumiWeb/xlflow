namespace Xlflow.ExcelBridge.Services;

public sealed record RunnerCommandArguments(
    string Action,
    string WorkbookPath,
    bool Visible);
