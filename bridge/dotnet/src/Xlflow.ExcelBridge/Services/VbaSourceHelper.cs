using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Source helper degrades file/hash failures into best-effort results.")]
internal static class VbaSourceHelper
{
    private static readonly Regex FolderAnnotationPattern = new(
        @"^'?@Folder\(\s*""([^""]*)""\s*\)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex AttributeVbPattern = new(
        @"^Attribute\s+VB_",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly JsonSerializerOptions FingerprintJsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    };

    private static readonly JsonSerializerOptions CompactJsonOptions = new()
    {
        WriteIndented = false,
    };

    public static bool IsSidecarMode(string? codeSource)
    {
        return string.Equals(codeSource?.Trim(), "sidecar", StringComparison.OrdinalIgnoreCase);
    }

    public static string? ParseFolderAnnotation(string text)
    {
        foreach (var line in SplitLines(text))
        {
            var match = FolderAnnotationPattern.Match(line.Trim());
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
        }
        return null;
    }

    public static string[] ParseFolderAnnotationSegments(string? annotation)
    {
        if (string.IsNullOrWhiteSpace(annotation))
        {
            return [];
        }

        var segments = annotation.Split('.', StringSplitOptions.RemoveEmptyEntries);
        var result = new List<string>();
        foreach (var segment in segments)
        {
            var cleaned = CleanFolderPathSegment(segment);
            if (!string.IsNullOrWhiteSpace(cleaned))
            {
                result.Add(cleaned);
            }
        }
        return result.ToArray();
    }

    public static string? BuildFolderAnnotation(string[] segments)
    {
        if (segments.Length == 0)
        {
            return null;
        }
        var clean = new List<string>();
        foreach (var segment in segments)
        {
            var cleaned = CleanFolderPathSegment(segment);
            if (!string.IsNullOrWhiteSpace(cleaned))
            {
                clean.Add(cleaned);
            }
        }
        if (clean.Count == 0)
        {
            return null;
        }
        return string.Join(".", clean);
    }

    public static string[] GetRelativePathSegments(string rootDir, string filePath)
    {
        if (string.IsNullOrWhiteSpace(rootDir) || string.IsNullOrWhiteSpace(filePath))
        {
            return [];
        }

        var rootFull = Path.GetFullPath(rootDir);
        var fileFull = Path.GetFullPath(filePath);
        var parentDir = Path.GetDirectoryName(fileFull) ?? "";

        if (!parentDir.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase))
        {
            return [];
        }

        var relative = Path.GetRelativePath(rootFull, parentDir);
        if (string.IsNullOrWhiteSpace(relative) || relative == ".")
        {
            return [];
        }

        var segments = new List<string>();
        foreach (var part in relative.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        {
            if (part == "..")
            {
                throw new InvalidOperationException($"path '{filePath}' resolves outside root '{rootDir}'");
            }
            var cleaned = CleanFolderPathSegment(part);
            if (!string.IsNullOrWhiteSpace(cleaned))
            {
                segments.Add(cleaned);
            }
        }
        return segments.ToArray();
    }

    public static string UpdateFolderAnnotationText(string text, string folderAnnotationMode, string? desiredAnnotation)
    {
        if (folderAnnotationMode is "ignore" or "preserve")
        {
            return text;
        }

        var lines = new List<string>(SplitLines(text));
        var (found, lineIndex) = FindFolderAnnotationLine(lines);

        string? annotationLine = null;
        if (!string.IsNullOrWhiteSpace(desiredAnnotation))
        {
            annotationLine = $"'@Folder(\"{desiredAnnotation}\")";
        }

        if (found)
        {
            if (annotationLine is null)
            {
                lines.RemoveAt(lineIndex);
            }
            else
            {
                lines[lineIndex] = annotationLine;
            }
        }
        else if (annotationLine is not null)
        {
            var insertIndex = GetFolderAnnotationInsertIndex(lines);
            lines.Insert(insertIndex, annotationLine);
        }

        return string.Join(Environment.NewLine, lines);
    }

    public static string NormalizeDocumentModuleContent(string text)
    {
        var lines = SplitLines(text);
        var filtered = new List<string>();
        var inClassHeader = false;
        var classHeaderBuffer = new List<string>();

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (trimmed == "VERSION 1.0 CLASS")
            {
                inClassHeader = true;
                classHeaderBuffer.Clear();
                classHeaderBuffer.Add(line);
                continue;
            }
            if (inClassHeader)
            {
                classHeaderBuffer.Add(line);
                if (trimmed == "END")
                {
                    inClassHeader = false;
                    classHeaderBuffer.Clear();
                }
                continue;
            }
            if (AttributeVbPattern.IsMatch(trimmed))
            {
                continue;
            }
            filtered.Add(line);
        }

        if (inClassHeader && classHeaderBuffer.Count > 0)
        {
            filtered.AddRange(classHeaderBuffer);
        }

        var hasOptionExplicit = false;
        var hasNonHeaderCode = false;
        foreach (var line in filtered)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                continue;
            }
            if (string.Equals(trimmed, "Option Explicit", StringComparison.OrdinalIgnoreCase))
            {
                hasOptionExplicit = true;
                continue;
            }
            hasNonHeaderCode = true;
        }

        if (!hasOptionExplicit && !hasNonHeaderCode)
        {
            filtered.Add("");
            filtered.Add("Option Explicit");
        }

        return string.Join(Environment.NewLine, filtered);
    }

    public static string GetCodeModuleText(object codeModule)
    {
        var lineCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(codeModule, "CountOfLines"));
        if (lineCount <= 0)
        {
            return "";
        }
        return ExcelBridgeSupport.GetString(codeModule, "Lines", 1, lineCount) ?? "";
    }

    public static void SetCodeModuleText(object codeModule, string text)
    {
        var lineCount = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(codeModule, "CountOfLines"));
        if (lineCount > 0)
        {
            ExcelBridgeSupport.InvokeViaDynamic(codeModule, "DeleteLines", 1, lineCount);
        }
        if (!string.IsNullOrWhiteSpace(text))
        {
            ExcelBridgeSupport.InvokeViaDynamic(codeModule, "InsertLines", 1, text);
        }
    }

    public static string? GetUserFormCodePath(string formsDir, string formName)
    {
        if (string.IsNullOrWhiteSpace(formsDir) || string.IsNullOrWhiteSpace(formName))
        {
            return null;
        }
        return Path.Combine(formsDir, "code", formName + ".bas");
    }

    public static string GetUserFormCodeDir(string formsDir)
    {
        return Path.Combine(formsDir, "code");
    }

    public static SourceFingerprint ComputeFingerprint(
        string workbookPath,
        string modulesDir,
        string classesDir,
        string formsDir,
        string workbookDir,
        string? codeSource)
    {
        var files = new List<SourceFileEntry>();
        foreach (var file in DiscoverSourceFiles(modulesDir, classesDir, formsDir, workbookDir, codeSource))
        {
            files.Add(new SourceFileEntry
            {
                Kind = file.Kind,
                Path = file.RelativePath,
                Hash = ComputeFileHash(file.FullPath),
            });
        }
        return new SourceFingerprint
        {
            WorkbookPath = Path.GetFullPath(workbookPath),
            Files = files.ToArray(),
        };
    }

    public static bool FingerprintMatchesState(SourceFingerprint fingerprint, string statePath)
    {
        if (string.IsNullOrWhiteSpace(statePath) || !File.Exists(statePath))
        {
            return false;
        }
        try
        {
            var existingJson = File.ReadAllText(statePath);
            var currentJson = JsonSerializer.Serialize(fingerprint, FingerprintJsonOptions);
            var existingNorm = NormalizeFingerprintJson(existingJson);
            var currentNorm = NormalizeFingerprintJson(currentJson);
            return existingNorm == currentNorm;
        }
        catch
        {
            return false;
        }
    }

    public static void WriteFingerprintState(SourceFingerprint fingerprint, string statePath)
    {
        if (string.IsNullOrWhiteSpace(statePath))
        {
            return;
        }
        var parent = Path.GetDirectoryName(statePath);
        if (!string.IsNullOrWhiteSpace(parent))
        {
            Directory.CreateDirectory(parent);
        }
        var json = JsonSerializer.Serialize(fingerprint, FingerprintJsonOptions);
        File.WriteAllText(statePath, json, new UTF8Encoding(false));
    }

    public static List<DiscoveredSourceFile> DiscoverSourceFiles(
        string modulesDir,
        string classesDir,
        string formsDir,
        string workbookDir,
        string? codeSource)
    {
        var files = new List<DiscoveredSourceFile>();

        AddFilesFromDir(files, modulesDir, "module", ".bas", null);
        AddFilesFromDir(files, classesDir, "class", ".cls", null);

        var formsCodeDir = IsSidecarMode(codeSource) ? GetUserFormCodeDir(formsDir) : null;
        AddFilesFromDir(files, formsDir, "form", ".bas", formsCodeDir);
        AddFilesFromDir(files, formsDir, "form", ".cls", formsCodeDir);
        AddFilesFromDir(files, formsDir, "form", ".frm", formsCodeDir);
        AddFilesFromDir(files, formsDir, "form", ".frx", formsCodeDir);

        AddFilesFromDir(files, workbookDir, "document", ".bas", null);

        if (IsSidecarMode(codeSource) && !string.IsNullOrWhiteSpace(formsDir))
        {
            var codeDir = GetUserFormCodeDir(formsDir);
            if (Directory.Exists(codeDir))
            {
                foreach (var file in Directory.GetFiles(codeDir, "*.bas", SearchOption.AllDirectories))
                {
                    files.Add(new DiscoveredSourceFile
                    {
                        Kind = "form_code",
                        RootDir = codeDir,
                        FullPath = file,
                        RelativePath = Path.GetRelativePath(codeDir, file).Replace('\\', '/'),
                        Extension = ".bas",
                        ModuleName = Path.GetFileNameWithoutExtension(file),
                    });
                }
            }
        }

        return files;
    }

    public static List<DuplicateModuleInfo> FindDuplicateModuleNames(List<DiscoveredSourceFile> files)
    {
        var seen = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var originalNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            if (file.Extension == ".frx" || file.Kind == "form_code")
            {
                continue;
            }
            var key = file.ModuleName.ToLowerInvariant();
            if (!seen.TryGetValue(key, out var paths))
            {
                paths = new List<string>();
                seen[key] = paths;
                originalNames[key] = file.ModuleName;
            }
            paths.Add(file.RelativePath);
        }

        var duplicates = new List<DuplicateModuleInfo>();
        foreach (var key in seen.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
        {
            if (seen[key].Count < 2)
            {
                continue;
            }
            duplicates.Add(new DuplicateModuleInfo
            {
                ModuleName = originalNames[key],
                Paths = seen[key].ToArray(),
            });
        }
        return duplicates;
    }

    public static string PrepareSourceForImport(string sourcePath, string destPath, string? rootDir, string folderAnnotationMode)
    {
        var parent = Path.GetDirectoryName(destPath);
        if (!string.IsNullOrWhiteSpace(parent))
        {
            Directory.CreateDirectory(parent);
        }

        var ext = Path.GetExtension(sourcePath).ToLowerInvariant();
        if (ext == ".frx")
        {
            File.Copy(sourcePath, destPath, true);
            return destPath;
        }

        var content = File.ReadAllText(sourcePath, Encoding.UTF8);
        if (!string.IsNullOrWhiteSpace(rootDir))
        {
            var desiredAnnotation = BuildFolderAnnotation(GetRelativePathSegments(rootDir, sourcePath));
            content = UpdateFolderAnnotationText(content, folderAnnotationMode, desiredAnnotation);
        }

        var encoding = GetVbaInteropEncoding();
        File.WriteAllText(destPath, content, encoding);
        return destPath;
    }

    public static string ReadExportedTextAsUtf8(string exportedFile)
    {
        var content = File.ReadAllText(exportedFile, GetVbaInteropEncoding());
        return ConvertToUtf8(content);
    }

    public static Encoding GetVbaInteropEncoding()
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(932);
        }
        catch
        {
            return new UTF8Encoding(false);
        }
    }

    public static Encoding GetVbaImportEncoding()
    {
        return GetVbaInteropEncoding();
    }

    public static string ConvertToUtf8(string content)
    {
        if (string.IsNullOrEmpty(content))
        {
            return content;
        }

        content = content.Replace("\r\n", "\n");
        content = content.Replace("\r", "\n");
        content = content.Replace("\n", "\r\n");

        if (content.Length > 0 && content[0] == '\uFEFF')
        {
            content = content[1..];
        }

        return content;
    }

    public static string GetFolderAnnotationForPath(string rootDir, string filePath)
    {
        var segments = GetRelativePathSegments(rootDir, filePath);
        return BuildFolderAnnotation(segments) ?? "";
    }

    public static string? GetComponentRootDir(int componentType, string modulesDir, string classesDir, string formsDir, string workbookDir)
    {
        return componentType switch
        {
            1 => modulesDir,
            2 => classesDir,
            3 => formsDir,
            100 => workbookDir,
            _ => null,
        };
    }

    public static string? GetComponentExtension(int componentType)
    {
        return componentType switch
        {
            1 => ".bas",
            2 => ".cls",
            3 => ".frm",
            100 => ".bas",
            _ => null,
        };
    }

    public static string? GetComponentPath(
        object component,
        string modulesDir,
        string classesDir,
        string formsDir,
        string workbookDir,
        bool folders,
        string folderAnnotation,
        bool defaultComponentFolders)
    {
        var type = ExcelBridgeSupport.ToInt(ExcelBridgeSupport.Get(component, "Type"));
        var name = ExcelBridgeSupport.GetString(component, "Name") ?? "";
        var rootDir = GetComponentRootDir(type, modulesDir, classesDir, formsDir, workbookDir);
        var extension = GetComponentExtension(type);

        if (string.IsNullOrWhiteSpace(rootDir) || string.IsNullOrWhiteSpace(extension))
        {
            return null;
        }

        var segments = Array.Empty<string>();
        if (folders && folderAnnotation != "ignore")
        {
            var annotation = GetFolderAnnotationForComponent(component, folderAnnotation);
            if (!string.IsNullOrWhiteSpace(annotation))
            {
                segments = ParseFolderAnnotationSegments(annotation);
            }
        }

        var path = rootDir;
        foreach (var segment in segments)
        {
            path = Path.Combine(path, segment);
        }
        return Path.Combine(path, name + extension);
    }

    private static string? GetFolderAnnotationForComponent(object component, string folderAnnotationMode)
    {
        if (folderAnnotationMode is "ignore" or "preserve")
        {
            return null;
        }

        try
        {
            object? codeModule = null;
            try
            {
                codeModule = ExcelBridgeSupport.Get(component, "CodeModule");
                var text = GetCodeModuleText(codeModule!);
                return ParseFolderAnnotation(text);
            }
            finally
            {
                ExcelBridgeSupport.ReleaseComObject(codeModule);
            }
        }
        catch
        {
            return null;
        }
    }

    private static void AddFilesFromDir(List<DiscoveredSourceFile> files, string dir, string kind, string extension, string? excludedDir)
    {
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
        {
            return;
        }

        var excludedFull = !string.IsNullOrWhiteSpace(excludedDir) ? Path.GetFullPath(excludedDir) : null;
        if (excludedFull is not null && !excludedFull.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
        {
            excludedFull += Path.DirectorySeparatorChar;
        }

        foreach (var file in Directory.GetFiles(dir, "*" + extension, SearchOption.AllDirectories))
        {
            var fileFull = Path.GetFullPath(file);
            if (excludedFull is not null && fileFull.StartsWith(excludedFull, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            files.Add(new DiscoveredSourceFile
            {
                Kind = kind,
                RootDir = dir,
                FullPath = file,
                RelativePath = Path.GetRelativePath(dir, file).Replace('\\', '/'),
                Extension = extension,
                ModuleName = Path.GetFileNameWithoutExtension(file),
            });
        }
    }

    private static string ComputeFileHash(string filePath)
    {
        try
        {
            var bytes = File.ReadAllBytes(filePath);
            var hash = SHA256.HashData(bytes);
            return Convert.ToHexString(hash).ToLowerInvariant();
        }
        catch
        {
            return "";
        }
    }

    private static string NormalizeFingerprintJson(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            return JsonSerializer.Serialize(doc.RootElement, CompactJsonOptions);
        }
        catch
        {
            return json;
        }
    }

    private static (bool Found, int LineIndex) FindFolderAnnotationLine(List<string> lines)
    {
        for (var i = 0; i < lines.Count; i++)
        {
            if (FolderAnnotationPattern.IsMatch(lines[i].Trim()))
            {
                return (true, i);
            }
        }
        return (false, -1);
    }

    private static int GetFolderAnnotationInsertIndex(List<string> lines)
    {
        for (var i = 0; i < lines.Count; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed == "VERSION 1.0 CLASS" || AttributeVbPattern.IsMatch(trimmed))
            {
                continue;
            }
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                continue;
            }
            return i;
        }
        return 0;
    }

    private static string[] SplitLines(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return [];
        }
        return text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
    }

    private static string CleanFolderPathSegment(string segment)
    {
        if (string.IsNullOrWhiteSpace(segment))
        {
            return "";
        }
        var cleaned = segment.Trim();
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            cleaned = cleaned.Replace(c, '_');
        }
        return cleaned;
    }
}

internal sealed record SourceFingerprint
{
    public string WorkbookPath { get; init; } = "";
    public SourceFileEntry[] Files { get; init; } = [];
}

internal sealed record SourceFileEntry
{
    public string Kind { get; init; } = "";
    public string Path { get; init; } = "";
    public string Hash { get; init; } = "";
}

internal sealed record DiscoveredSourceFile
{
    public string Kind { get; init; } = "";
    public string RootDir { get; init; } = "";
    public string FullPath { get; init; } = "";
    public string RelativePath { get; init; } = "";
    public string Extension { get; init; } = "";
    public string ModuleName { get; init; } = "";
}

internal sealed record DuplicateModuleInfo
{
    public string ModuleName { get; init; } = "";
    public string[] Paths { get; init; } = [];
}
