namespace Xlflow.ExcelBridge.Services;

public sealed record TestCommandArguments(
    string WorkbookPath,
    string Filter,
    string ModuleFilter,
    string TagFilter,
    bool Visible,
    string RuntimeMode,
    string RuntimeSource,
    string MsgBoxResponsesJSON,
    string InputResponsesJSON,
    string FileDialogResponsesJSON,
    bool DebugStreamEnabled,
    string DebugStreamPipeName,
    bool UIStreamEnabled,
    string UIStreamPipeName,
    bool UIStreamRedactInput,
    bool UseSession,
    string MetadataPath);
