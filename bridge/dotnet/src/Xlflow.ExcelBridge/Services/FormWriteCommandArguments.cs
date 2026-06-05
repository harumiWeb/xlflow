namespace Xlflow.ExcelBridge.Services;

public sealed record FormWriteCommandArguments(
    string Action,
    string WorkbookPath,
    string SpecPath,
    string FormsDir,
    string CodeSource,
    bool Folders,
    string FolderAnnotation,
    bool DefaultComponentFolders,
    string SpecJson64,
    bool Overwrite,
    bool NoSave,
    bool Visible,
    bool UseSession,
    string MetadataPath);
