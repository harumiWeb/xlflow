namespace Xlflow.ExcelBridge.Services;

public sealed record ProcessCommandArguments(
    string Action,
    int? TargetPid,
    bool Auto,
    bool All);
