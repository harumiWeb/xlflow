namespace Xlflow.ExcelBridge.Services;

public sealed record PushCommandArguments(
    string WorkbookPath,
    string ModulesDir,
    string ClassesDir,
    string FormsDir,
    string WorkbookDir,
    string CodeSource,
    string BackupRoot,
    bool Folders,
    string FolderAnnotation,
    bool DefaultComponentFolders,
    string StatePath,
    bool Visible,
    string BackupMode,
    bool ChangedOnly,
    bool UseSession,
    bool NoSave,
    string MetadataPath);
