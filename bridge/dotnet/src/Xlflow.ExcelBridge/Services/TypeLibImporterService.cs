using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using Microsoft.Win32;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "TypeLib import is a Windows-only bridge command.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Importer failures are normalized into structured bridge errors.")]
[SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "Command services keep instance Execute methods for consistency with other bridge services.")]
public sealed class TypeLibImporterService
{
    private const int MemberIDNil = -1;
    private const int DispIDValue = 0;
    private const short VariantTypeUserDefined = 29;
    private const short VariantTypePtr = 26;
    private const short VariantTypeSafeArray = 27;
    private const short FuncFlagRestricted = 0x1;
    private const short FuncFlagHidden = 0x40;
    private const short VarFlagRestricted = 0x1;
    private const short VarFlagHidden = 0x40;

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private static readonly Dictionary<string, LibraryTarget> KnownLibraries = new(StringComparer.OrdinalIgnoreCase)
    {
        ["excel"] = new("Excel", "{00020813-0000-0000-C000-000000000046}", "excel.generated.json"),
        ["office"] = new("Office", "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "office.generated.json"),
        ["msforms"] = new("MSForms", "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", "msforms.generated.json"),
        ["scripting"] = new("Scripting", "{420B2830-E718-11CF-893D-00A0C9054228}", "scripting.generated.json"),
        ["adodb"] = new("ADODB", "{00000205-0000-0010-8000-00AA006D2EA4}", "adodb.generated.json"),
        ["vbide"] = new("VBIDE", "{0002E157-0000-0000-C000-000000000046}", "vbide.generated.json"),
    };

    public BridgeResponse Execute(BridgeRequest request, TypeDbImportArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(args.OutputDir))
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "type_db_args_invalid",
                Message: "OutputDir is required",
                Phase: "type-db-import",
                Source: "xlflow"));
        }

        try
        {
            var outputDir = Path.GetFullPath(args.OutputDir);
            Directory.CreateDirectory(outputDir);
            var libraries = ResolveLibraries(args.Libraries);
            RemoveStaleGeneratedFiles(outputDir);
            var manifestLibraries = new List<Dictionary<string, object?>>();
            var generatedFiles = new List<string>();
            var logs = new List<string>();

            foreach (var target in libraries.Targets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                TypeLibRegistration registration;
                try
                {
                    registration = ResolveRegisteredTypeLib(target);
                }
                catch (InvalidOperationException ex) when (libraries.BestEffort)
                {
                    logs.Add($"skipped {target.Name}: {ex.Message}");
                    continue;
                }
                Dictionary<string, object?> db;
                try
                {
                    db = ImportLibrary(target, registration);
                }
                catch (Exception ex) when (libraries.BestEffort)
                {
                    logs.Add($"skipped {target.Name}: {ex.Message}");
                    continue;
                }
                var outputPath = Path.Combine(outputDir, target.Output);
                File.WriteAllText(outputPath, JsonSerializer.Serialize(db, JsonOptions) + Environment.NewLine);
                generatedFiles.Add(outputPath);
                manifestLibraries.Add(new Dictionary<string, object?>
                {
                    ["name"] = target.Name,
                    ["libid"] = target.LibID,
                    ["major"] = registration.Major,
                    ["minor"] = registration.Minor,
                    ["lcid"] = registration.LCID,
                    ["source"] = "registry",
                    ["output"] = target.Output,
                });
            }

            if (generatedFiles.Count == 0)
            {
                throw new InvalidOperationException("No requested TypeLib libraries were found on this machine.");
            }

            var manifest = new Dictionary<string, object?>
            {
                ["schema_version"] = 1,
                ["generator"] = "xlflow",
                ["generator_version"] = string.IsNullOrWhiteSpace(args.GeneratorVersion) ? "dev" : args.GeneratorVersion,
                ["generated_at"] = DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture),
                ["platform"] = "windows",
                ["arch"] = RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
                ["libraries"] = manifestLibraries,
            };
            File.WriteAllText(Path.Combine(outputDir, "manifest.json"), JsonSerializer.Serialize(manifest, JsonOptions) + Environment.NewLine);

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = [$"generated {generatedFiles.Count} TypeLib type database file(s)", .. logs],
                Extensions = new Dictionary<string, object?>
                {
                    ["type_db"] = new Dictionary<string, object?>
                    {
                        ["dir"] = outputDir,
                        ["manifest_path"] = Path.Combine(outputDir, "manifest.json"),
                        ["generated_files"] = generatedFiles,
                        ["libraries"] = manifestLibraries,
                    },
                },
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "type_db_import_failed",
                Message: ex.Message,
                Phase: "type-db-import",
                Source: "xlflow-excel-bridge"));
        }
    }

    private static void RemoveStaleGeneratedFiles(string outputDir)
    {
        foreach (var path in Directory.EnumerateFiles(outputDir, "*.generated.json"))
        {
            File.Delete(path);
        }
        var manifestPath = Path.Combine(outputDir, "manifest.json");
        if (File.Exists(manifestPath))
        {
            File.Delete(manifestPath);
        }
    }

    private static ResolvedLibraries ResolveLibraries(string libraries)
    {
        var names = libraries.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (names.Length == 0)
        {
            names = ["excel"];
        }
        if (names.Any(name => string.Equals(name, "all", StringComparison.OrdinalIgnoreCase)))
        {
            return new ResolvedLibraries(KnownLibraries.Values.ToList(), BestEffort: true);
        }
        var targets = new List<LibraryTarget>();
        foreach (var name in names)
        {
            if (!KnownLibraries.TryGetValue(name, out var target))
            {
                throw new InvalidOperationException("Unknown TypeLib library: " + name);
            }
            if (!targets.Any(item => string.Equals(item.Name, target.Name, StringComparison.OrdinalIgnoreCase)))
            {
                targets.Add(target);
            }
        }
        return new ResolvedLibraries(targets, BestEffort: false);
    }

    private static TypeLibRegistration ResolveRegisteredTypeLib(LibraryTarget target)
    {
        using var root = Registry.ClassesRoot.OpenSubKey(@"TypeLib\" + target.LibID);
        if (root is null)
        {
            throw new InvalidOperationException(target.Name + " TypeLib was not found on this machine.");
        }
        var version = root.GetSubKeyNames()
            .Select(ParseVersion)
            .Where(item => item is not null)
            .Select(item => item!.Value)
            .OrderByDescending(item => item.Major)
            .ThenByDescending(item => item.Minor)
            .FirstOrDefault();
        if (version == default)
        {
            throw new InvalidOperationException(target.Name + " TypeLib version was not found on this machine.");
        }
        return new TypeLibRegistration(target.LibID, version.Major, version.Minor, 0);
    }

    private static (int Major, int Minor)? ParseVersion(string value)
    {
        var parts = value.Split('.', 2);
        if (parts.Length != 2 ||
            !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var major) ||
            !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var minor))
        {
            return null;
        }
        return (major, minor);
    }

    private static Dictionary<string, object?> ImportLibrary(LibraryTarget target, TypeLibRegistration registration)
    {
        var guid = Guid.Parse(registration.LibID);
        LoadRegTypeLib(ref guid, (ushort)registration.Major, (ushort)registration.Minor, registration.LCID, out var typeLib);
        var types = new List<Dictionary<string, object?>>();
        var constants = new List<Dictionary<string, object?>>();
        var classIDs = new Dictionary<Guid, string>();
        var count = typeLib.GetTypeInfoCount();
        for (var i = 0; i < count; i++)
        {
            typeLib.GetTypeInfo(i, out var typeInfo);
            typeInfo.GetTypeAttr(out var typeAttrPtr);
            try
            {
                var attr = Marshal.PtrToStructure<TYPEATTR>(typeAttrPtr);
                typeInfo.GetDocumentation(MemberIDNil, out var rawName, out var summary, out _, out _);
                var name = CanonicalTypeName(target.Name, rawName);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }
                if (attr.typekind == TYPEKIND.TKIND_COCLASS && attr.guid != Guid.Empty)
                {
                    classIDs[attr.guid] = name;
                }
                if (attr.typekind == TYPEKIND.TKIND_ENUM)
                {
                    constants.AddRange(ImportConstants(typeInfo, attr, target.Name, name));
                    types.Add(new Dictionary<string, object?>
                    {
                        ["name"] = name,
                        ["library"] = target.Name,
                        ["kind"] = "enum",
                        ["summary"] = EmptyToNull(summary),
                        ["confidence"] = "generated",
                        ["source"] = "typelib",
                    });
                    continue;
                }

                var type = new Dictionary<string, object?>
                {
                    ["name"] = name,
                    ["library"] = target.Name,
                    ["kind"] = TypeKind(attr.typekind),
                    ["summary"] = EmptyToNull(summary),
                    ["confidence"] = "generated",
                    ["source"] = "typelib",
                };
                var members = ImportMembers(typeInfo, attr, target.Name);
                if (members.Properties.Count > 0)
                {
                    type["properties"] = members.Properties;
                }
                if (members.Methods.Count > 0)
                {
                    type["methods"] = members.Methods;
                }
                types.Add(type);
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(typeAttrPtr);
            }
        }

        var db = new Dictionary<string, object?>
        {
            ["types"] = types.OrderBy(item => item["name"]).ToArray(),
            ["constants"] = constants.OrderBy(item => item["name"]).ToArray(),
        };
        var progIDs = DiscoverProgIDs(target.LibID, classIDs);
        if (progIDs.Count > 0)
        {
            db["progids"] = progIDs;
        }
        return db;
    }

    private static Dictionary<string, string> DiscoverProgIDs(string libID, IReadOnlyDictionary<Guid, string> classIDs)
    {
        return SelectProgIDsForTypeLib(libID, classIDs, EnumerateRegisteredProgIDs());
    }

    internal static Dictionary<string, string> SelectProgIDsForTypeLib(
        string libID,
        IReadOnlyDictionary<Guid, string> classIDs,
        IEnumerable<RegisteredProgID> registrations)
    {
        var outMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var targetLibID = NormalizeGuid(libID);
        foreach (var registration in registrations)
        {
            if (!string.IsNullOrWhiteSpace(registration.TypeLib) &&
                !string.Equals(NormalizeGuid(registration.TypeLib), targetLibID, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            if (!classIDs.TryGetValue(registration.ClassID, out var typeName) || string.IsNullOrWhiteSpace(typeName))
            {
                continue;
            }
            foreach (var progID in registration.ProgIDs)
            {
                if (!string.IsNullOrWhiteSpace(progID))
                {
                    outMap[progID] = typeName;
                }
            }
        }
        return outMap
            .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<RegisteredProgID> EnumerateRegisteredProgIDs()
    {
        using var clsidRoot = Registry.ClassesRoot.OpenSubKey("CLSID");
        if (clsidRoot is null)
        {
            yield break;
        }
        foreach (var classIDName in clsidRoot.GetSubKeyNames())
        {
            if (!Guid.TryParse(classIDName, out var classID))
            {
                continue;
            }
            using var classKey = clsidRoot.OpenSubKey(classIDName);
            if (classKey is null)
            {
                continue;
            }
            var typeLib = StringValue(classKey.OpenSubKey("TypeLib"));
            var progIDs = new[]
                {
                    StringValue(classKey.OpenSubKey("VersionIndependentProgID")),
                    StringValue(classKey.OpenSubKey("ProgID")),
                }
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (progIDs.Length == 0)
            {
                continue;
            }
            yield return new RegisteredProgID(classID, typeLib, progIDs);
        }
    }

    private static string StringValue(RegistryKey? key)
    {
        using (key)
        {
            return key?.GetValue(null) as string ?? "";
        }
    }

    private static string NormalizeGuid(string value)
    {
        return Guid.TryParse(value, out var guid) ? guid.ToString("B").ToUpperInvariant() : value.Trim().ToUpperInvariant();
    }

    private static ImportedMembers ImportMembers(ITypeInfo typeInfo, TYPEATTR attr, string library)
    {
        var properties = new Dictionary<string, Dictionary<string, object?>>(StringComparer.OrdinalIgnoreCase);
        var propertiesWithGetter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var methods = new Dictionary<string, Dictionary<string, object?>>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < attr.cFuncs; i++)
        {
            typeInfo.GetFuncDesc(i, out var funcDescPtr);
            try
            {
                var desc = Marshal.PtrToStructure<FUNCDESC>(funcDescPtr);
                if ((desc.wFuncFlags & (FuncFlagRestricted | FuncFlagHidden)) != 0)
                {
                    continue;
                }
                var names = GetNames(typeInfo, desc.memid, desc.cParams + 1);
                var memberName = names.Length > 0 ? names[0] : "";
                if (string.IsNullOrWhiteSpace(memberName))
                {
                    continue;
                }
                var member = new Dictionary<string, object?>
                {
                    ["name"] = memberName,
                    ["return_type"] = TypeName(typeInfo, desc.elemdescFunc.tdesc, library),
                    ["parameters"] = Parameters(typeInfo, desc, names, library),
                };
                if (desc.memid == DispIDValue)
                {
                    member["default"] = true;
                }
                if (desc.invkind is INVOKEKIND.INVOKE_PROPERTYGET or INVOKEKIND.INVOKE_PROPERTYPUT or INVOKEKIND.INVOKE_PROPERTYPUTREF)
                {
                    if (!properties.TryGetValue(memberName, out var property))
                    {
                        property = new Dictionary<string, object?>
                        {
                            ["name"] = memberName,
                        };
                        properties[memberName] = property;
                    }
                    if (desc.memid == DispIDValue)
                    {
                        property["default"] = true;
                    }
                    if (desc.invkind == INVOKEKIND.INVOKE_PROPERTYGET)
                    {
                        property["return_type"] = member["return_type"];
                        property["parameters"] = member["parameters"];
                        property.Remove("write_only");
                        propertiesWithGetter.Add(memberName);
                    }
                    else if (!propertiesWithGetter.Contains(memberName))
                    {
                        property["write_only"] = true;
                    }
                    continue;
                }
                methods[memberName] = member;
            }
            finally
            {
                typeInfo.ReleaseFuncDesc(funcDescPtr);
            }
        }
        return new ImportedMembers(properties.Values.ToList(), methods.Values.ToList());
    }

    private static List<Dictionary<string, object?>> ImportConstants(ITypeInfo typeInfo, TYPEATTR attr, string library, string enumName)
    {
        var constants = new List<Dictionary<string, object?>>();
        for (var i = 0; i < attr.cVars; i++)
        {
            typeInfo.GetVarDesc(i, out var varDescPtr);
            try
            {
                var desc = Marshal.PtrToStructure<VARDESC>(varDescPtr);
                if ((desc.wVarFlags & (VarFlagRestricted | VarFlagHidden)) != 0)
                {
                    continue;
                }
                typeInfo.GetDocumentation(desc.memid, out var name, out var summary, out _, out _);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }
                constants.Add(new Dictionary<string, object?>
                {
                    ["name"] = name,
                    ["library"] = library,
                    ["kind"] = "constant",
                    ["type"] = TypeName(typeInfo, desc.elemdescVar.tdesc, library),
                    ["value"] = ConstantValue(desc),
                    ["enum_group"] = enumName.Contains('.') ? enumName[(enumName.LastIndexOf('.') + 1)..] : enumName,
                    ["summary"] = EmptyToNull(summary),
                });
            }
            finally
            {
                typeInfo.ReleaseVarDesc(varDescPtr);
            }
        }
        return constants;
    }

    private static List<Dictionary<string, object?>> Parameters(ITypeInfo typeInfo, FUNCDESC desc, string[] names, string library)
    {
        var parameters = new List<Dictionary<string, object?>>();
        var elemSize = Marshal.SizeOf<ELEMDESC>();
        for (var i = 0; i < desc.cParams; i++)
        {
            var ptr = IntPtr.Add(desc.lprgelemdescParam, i * elemSize);
            var elem = Marshal.PtrToStructure<ELEMDESC>(ptr);
            var paramName = names.Length > i + 1 && !string.IsNullOrWhiteSpace(names[i + 1])
                ? names[i + 1]
                : "Arg" + (i + 1).ToString(CultureInfo.InvariantCulture);
            var isParamArray = desc.cParamsOpt < 0 && i == desc.cParams - 1;
            var optionalByPosition = desc.cParamsOpt > 0 && i >= desc.cParams - desc.cParamsOpt;
            parameters.Add(new Dictionary<string, object?>
            {
                ["name"] = paramName,
                ["type"] = TypeName(typeInfo, elem.tdesc, library),
                ["optional"] = isParamArray || optionalByPosition || (elem.desc.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FOPT) != 0,
                ["param_array"] = isParamArray ? true : null,
            });
        }
        return parameters;
    }

    private static string[] GetNames(ITypeInfo typeInfo, int memid, int count)
    {
        var names = new string[Math.Max(1, count)];
        typeInfo.GetNames(memid, names, names.Length, out var actual);
        return names.Take(actual).Where(name => !string.IsNullOrWhiteSpace(name)).ToArray();
    }

    private static string TypeName(ITypeInfo typeInfo, TYPEDESC desc, string library)
    {
        var vt = desc.vt;
        if (vt == VariantTypePtr || vt == VariantTypeSafeArray)
        {
            var nested = Marshal.PtrToStructure<TYPEDESC>(desc.lpValue);
            return TypeName(typeInfo, nested, library);
        }
        if (vt == VariantTypeUserDefined)
        {
            var href = unchecked((int)desc.lpValue.ToInt64());
            typeInfo.GetRefTypeInfo(href, out var refTypeInfo);
            refTypeInfo.GetDocumentation(MemberIDNil, out var rawName, out _, out _, out _);
            return CanonicalTypeName(library, rawName);
        }
        return vt switch
        {
            0 => "",
            1 => "Null",
            2 => "Integer",
            3 => "Long",
            4 => "Single",
            5 => "Double",
            6 => "Currency",
            7 => "Date",
            8 => "String",
            9 => "Object",
            10 => "Error",
            11 => "Boolean",
            12 => "Variant",
            13 => "Object",
            14 => "Decimal",
            16 => "Byte",
            17 => "Byte",
            18 => "Integer",
            19 => "Long",
            20 => "LongLong",
            21 => "LongLong",
            22 => "Long",
            23 => "Long",
            24 => "",
            25 => "HRESULT",
            _ => "Variant",
        };
    }

    private static string CanonicalTypeName(string library, string rawName)
    {
        rawName = rawName.Trim();
        while (rawName.StartsWith('_'))
        {
            rawName = rawName[1..];
        }
        if (rawName.Length == 0)
        {
            return "";
        }
        return rawName.Contains('.') ? rawName : library + "." + rawName;
    }

    private static string TypeKind(TYPEKIND kind)
    {
        return kind switch
        {
            TYPEKIND.TKIND_DISPATCH => "interface",
            TYPEKIND.TKIND_INTERFACE => "interface",
            TYPEKIND.TKIND_COCLASS => "class",
            TYPEKIND.TKIND_RECORD => "record",
            TYPEKIND.TKIND_ALIAS => "alias",
            TYPEKIND.TKIND_MODULE => "module",
            TYPEKIND.TKIND_UNION => "union",
            _ => "type",
        };
    }

    private static string? EmptyToNull(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string ConstantValue(VARDESC desc)
    {
        if (desc.varkind != VARKIND.VAR_CONST || desc.desc.lpvarValue == IntPtr.Zero)
        {
            return "";
        }
        try
        {
            var value = Marshal.GetObjectForNativeVariant(desc.desc.lpvarValue);
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
        catch
        {
            return "";
        }
    }

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void LoadRegTypeLib(ref Guid rguid, ushort wVerMajor, ushort wVerMinor, int lcid, out ITypeLib typeLib);

    private sealed record LibraryTarget(string Name, string LibID, string Output);

    private sealed record ResolvedLibraries(List<LibraryTarget> Targets, bool BestEffort);

    private sealed record TypeLibRegistration(string LibID, int Major, int Minor, int LCID);

    private sealed record ImportedMembers(List<Dictionary<string, object?>> Properties, List<Dictionary<string, object?>> Methods);
}

public sealed record TypeDbImportArguments(string OutputDir, string GeneratorVersion, string Libraries);

internal sealed record RegisteredProgID(Guid ClassID, string TypeLib, IReadOnlyList<string> ProgIDs);
