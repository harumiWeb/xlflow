namespace Xlflow.ExcelBridge.Services;

public sealed record PullCommandArguments(
    string WorkbookPath,
    string ModulesDir,
    string ClassesDir,
    string FormsDir,
    string WorkbookDir,
    string CodeSource,
    bool Folders,
    string FolderAnnotation,
    bool DefaultComponentFolders,
    bool Visible,
    bool UseSession,
    string MetadataPath);
