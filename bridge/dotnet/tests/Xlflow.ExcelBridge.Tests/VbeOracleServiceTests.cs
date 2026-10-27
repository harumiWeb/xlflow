using System.Text;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class VbeOracleServiceTests
{
    [Fact]
    public void SchemaV1AcceptsClassFormAndDocumentBasBodies()
    {
        var plan = new
        {
            schema_version = 1,
            case_id = "module-kinds",
            probe_mode = "compile",
            modules = new object[]
            {
                new { name = "AClass", kind = "class", source = "Attribute VB_Name = \"AClass\"\nPublic Sub Run()\nEnd Sub\n" },
                new { name = "AForm", kind = "form", source = "Attribute VB_Name = \"AForm\"\nPublic Sub Run()\nEnd Sub\n" },
                new { name = "ThisWorkbook", kind = "document", document_target = "workbook", source = "Attribute VB_Name = \"ThisWorkbook\"\nPrivate Sub Workbook_Open()\nEnd Sub\n" },
                new { name = "Sheet1", kind = "document", document_target = "worksheet", source = "Attribute VB_Name = \"Sheet1\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n" },
            },
        };
        AssertPlanAccepted(plan);
    }

    [Fact]
    public void SchemaV1RejectsDesignerArtifacts()
    {
        var plan = new
        {
            schema_version = 1,
            case_id = "designer-artifact",
            probe_mode = "compile",
            modules = new[] { new { name = "Form1", kind = "form", source_path = "Form1.frm" } },
        };
        var response = Execute(plan);

        Assert.Equal("vbe_oracle_plan_invalid", response.Error?.Code);
    }

    [Fact]
    public void SchemaV1RejectsNonBasPathEvenForInlineSource()
    {
        var plan = new
        {
            schema_version = 1,
            case_id = "non-bas-inline",
            probe_mode = "compile",
            modules = new[]
            {
                new { name = "Class1", kind = "class", source_path = "Class1.cls", source = "Attribute VB_Name = \"Class1\"\n" },
            },
        };
        var response = Execute(plan);

        Assert.Equal("vbe_oracle_plan_invalid", response.Error?.Code);
    }

    [Fact]
    public void SchemaV1AcceptsExternalTypeLibReference()
    {
        var plan = new
        {
            schema_version = 1,
            case_id = "external-reference",
            probe_mode = "compile",
            references = new[]
            {
                new { name = "Outlook", libid = "{00062FFF-0000-0000-C000-000000000046}", major = 9, minor = 6 },
            },
            modules = new[]
            {
                new { name = "Main", kind = "standard", source = "Attribute VB_Name = \"Main\"\n" },
            },
        };

        AssertPlanAccepted(plan);
    }

    [Fact]
    public void OracleProvisioningAddsExternalTypeLibReference()
    {
        var plan = new
        {
            schema_version = 1,
            case_id = "external-reference",
            probe_mode = "compile",
            references = new[]
            {
                new { name = "Outlook", libid = "{00062FFF-0000-0000-C000-000000000046}", major = 9, minor = 6 },
            },
            modules = new[]
            {
                new { name = "Main", kind = "standard", source = "Attribute VB_Name = \"Main\"\n" },
            },
        };
        var json = JsonSerializer.Serialize(plan);
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        var serviceType = typeof(VbeOracleService);
        var decode = serviceType.GetMethod("DecodePlan", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var provision = serviceType.GetMethod("ProvisionReferences", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var decoded = decode.Invoke(null, [encoded])!;
        var references = decoded.GetType().GetProperty("References", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!.GetValue(decoded)!;
        var project = new FakeProject();

        provision.Invoke(null, [new FakeWorkbook(project), references, CancellationToken.None]);

        Assert.Equal(("{00062FFF-0000-0000-C000-000000000046}", 9, 6), project.References.Added.Single());
    }

    [Theory]
    [InlineData("provision_references")]
    [InlineData("provision_components")]
    public void ReferenceAndComponentProvisioningFailuresUseImportErrorCode(string stage)
    {
        var error = VbeOracleService.ClassifyInfrastructureCode(stage, new InvalidOperationException("provisioning failed"));

        Assert.Equal("oracle_import_failed", error);
    }

    [Fact]
    public void OracleSanitizerRemovesExportHeadersBeforeInjection()
    {
        var sanitizer = typeof(VbeOracleService).GetMethod("SanitizeOracleSource", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var result = (string)sanitizer.Invoke(null, ["Attribute VB_Name = \"Class1\"\nVERSION 1.0 CLASS\nBEGIN\nEND\nOption Explicit\n"])!;

        Assert.DoesNotContain("Attribute VB_Name", result);
        Assert.DoesNotContain("VERSION 1.0 CLASS", result);
        Assert.Contains("Option Explicit", result);
    }

    [Fact]
    public void OracleProvisioningRejectsVbeSourceMutation()
    {
        var setter = typeof(VbeOracleService).GetMethod("SetOracleCodeModuleText", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var module = new MutatingCodeModule();

        var error = Assert.Throws<System.Reflection.TargetInvocationException>(() =>
            setter.Invoke(null, [module, "Option Explicit\nPublic Sub Run()\nEnd Sub\n", "Main"]));

        var mutation = Assert.IsType<VbeOracleService.OracleSourceMutationException>(error.InnerException);
        Assert.Contains("changed oracle fixture source", mutation.Message);
        Assert.Equal("oracle_source_mutated", VbeOracleService.ClassifyInfrastructureCode("provision_components", mutation));
    }

    [Theory]
    [InlineData("True")]
    [InlineData("False")]
    public void OracleClassImportPreservesPredeclaredAttributeAndUsesCrLf(string value)
    {
        var source = "Attribute VB_Name = \"Widget\"\nAttribute VB_PredeclaredId = " + value + "\nOption Explicit\n";
        var text = VbeOracleService.BuildOracleClassImport(source);

        Assert.StartsWith("VERSION 1.0 CLASS\r\nBEGIN\r\n", text);
        Assert.Contains("Attribute VB_PredeclaredId = " + value + "\r\n", text);
        Assert.DoesNotContain("\n", text.Replace("\r\n", "", StringComparison.Ordinal));
        Assert.Equal(text, VbeOracleService.BuildOracleClassImport(text));
        VbeOracleService.VerifyOraclePredeclaredId(source, text);
    }

    [Theory]
    [InlineData("Attribute VB_PredeclaredId = False\r\n")]
    [InlineData("Option Explicit\r\n")]
    public void OracleClassImportRejectsLostPredeclaredAttribute(string exported)
    {
        Assert.Throws<VbeOracleService.OracleSourceMutationException>(() =>
            VbeOracleService.VerifyOraclePredeclaredId("Attribute VB_PredeclaredId = True\n", exported));
    }

    [Theory]
    [InlineData(true, "\n")]
    [InlineData(false, "\n")]
    [InlineData(false, "\r\n")]
    [InlineData(false, "\r")]
    public void OracleClassImportRelocatesExistingHeaderWithoutChangingSettings(bool headerFirst, string newline)
    {
        const string header = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = 0  'False\nEND\n";
        const string attributes = "Attribute VB_Name = \"Widget\"\nAttribute VB_PredeclaredId = True\n";
        var source = "\uFEFF" + (headerFirst ? header + attributes : attributes + header) + "Option Explicit\n";
        source = source.Replace("\n", newline, StringComparison.Ordinal);

        var text = VbeOracleService.BuildOracleClassImport(source);

        Assert.StartsWith(header.Replace("\n", "\r\n", StringComparison.Ordinal), text);
        Assert.Equal(2, text.Split("VERSION 1.0 CLASS", StringSplitOptions.None).Length);
        Assert.EndsWith(attributes.Replace("\n", "\r\n", StringComparison.Ordinal) + "Option Explicit\r\n", text);
        Assert.DoesNotContain("\n", text.Replace("\r\n", "", StringComparison.Ordinal));
        Assert.Equal(text, VbeOracleService.BuildOracleClassImport(text));
    }

    [Theory]
    [InlineData("True", true)]
    [InlineData("False", true)]
    [InlineData(null, false)]
    public void OracleProvisioningImportsExplicitPredeclaredClassesAndAddsOrdinaryClasses(string? value, bool import)
    {
        var source = "Attribute VB_Name = \"Widget\"\n";
        if (value is not null)
        {
            source += "Attribute VB_PredeclaredId = " + value + "\n";
        }
        source += "Option Explicit\nPublic Value As Long\n";
        var plan = new
        {
            schema_version = 1,
            case_id = "class-provisioning",
            probe_mode = "compile",
            modules = new[] { new { name = "Widget", kind = "class", source } },
        };
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(plan)));
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var decoded = typeof(VbeOracleService).GetMethod("DecodePlan", flags)!.Invoke(null, [encoded])!;
        var modules = decoded.GetType().GetProperty("Modules")!.GetValue(decoded)!;
        var project = new FakeProject();
        var tempDirectory = Directory.CreateTempSubdirectory("xlflow-oracle-components-").FullName;
        try
        {
            typeof(VbeOracleService).GetMethod("ProvisionComponents", flags)!
                .Invoke(null, [new FakeWorkbook(project), modules, tempDirectory, CancellationToken.None]);

            Assert.Equal(import ? 1 : 0, project.VBComponents.ImportCalls);
            Assert.Equal(import ? 0 : 1, project.VBComponents.AddCalls);
            var component = Assert.IsType<FakeComponent>(project.VBComponents.Component);
            Assert.Equal("Widget", component.Name);
            Assert.Equal(2, component.Type);
            Assert.Equal(VbaSourceHelper.NormalizeDocumentModuleContent(source), component.CodeModule.Text);
            Assert.Equal(import ? 1 : 0, component.ExportCalls);
            if (import)
            {
                Assert.Equal(".cls", Path.GetExtension(project.VBComponents.ImportPath));
                Assert.Contains("Attribute VB_PredeclaredId = " + value, component.ImportedSource);
            }
        }
        finally
        {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Theory]
    [InlineData(100, "ThisWorkbook", "ThisWorkbook", "workbook", true)]
    [InlineData(100, "THISWORKBOOK", "ThisWorkbook", "workbook", true)]
    [InlineData(100, "Sheet2", "Sheet2", "worksheet", true)]
    [InlineData(100, "Sheet1", "Sheet2", "worksheet", false)]
    [InlineData(100, "ThisWorkbook", "Sheet2", "worksheet", false)]
    [InlineData(2, "Sheet2", "Sheet2", "worksheet", false)]
    public void DocumentTargetResolutionRequiresMatchingModuleName(int componentType, string candidateName, string moduleName, string target, bool expected)
    {
        var resolver = typeof(VbeOracleService).GetMethod("MatchesDocumentTarget", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var result = (bool)resolver.Invoke(null, [componentType, candidateName, moduleName, target])!;

        Assert.Equal(expected, result);
    }

    private static BridgeResponse Execute(object plan)
    {
        var json = JsonSerializer.Serialize(plan);
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        return new VbeOracleService().Execute(
            new BridgeRequest
            {
                ProtocolVersion = ProtocolVersion.Current,
                RequestId = "req-oracle-service",
                Command = "vbe-oracle",
            },
            new VbeOracleCommandArguments(encoded, TimeSpan.FromSeconds(1)),
            CancellationToken.None);
    }

    private static void AssertPlanAccepted(object plan)
    {
        var json = JsonSerializer.Serialize(plan);
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        var serviceType = typeof(VbeOracleService);
        var decode = serviceType.GetMethod("DecodePlan", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var validate = serviceType.GetMethod("ValidatePlan", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var decoded = decode.Invoke(null, [encoded]);
        var exception = Record.Exception(() => validate.Invoke(null, [decoded]));
        Assert.Null(exception);
    }

    public sealed class MutatingCodeModule
    {
        private string text = "";

        public int CountOfLines => string.IsNullOrEmpty(text) ? 0 : text.Split('\n').Length;

        public string Lines(object startLine, object count) => text;

        public object? DeleteLines(object startLine, object count)
        {
            text = "";
            return null;
        }

        public object? InsertLines(object startLine, object value)
        {
            text = Convert.ToString(value) + "' VBE mutation";
            return null;
        }
    }

    private sealed class FakeWorkbook(FakeProject project)
    {
        public FakeProject VBProject { get; } = project;
    }

    private sealed class FakeProject
    {
        public FakeReferences References { get; } = new();
        public FakeComponents VBComponents { get; } = new();
    }

    private sealed class FakeComponents
    {
        public int ImportCalls { get; private set; }
        public int AddCalls { get; private set; }
        public string ImportPath { get; private set; } = "";
        public FakeComponent? Component { get; private set; }

        public FakeComponent Import(string path)
        {
            ImportCalls++;
            ImportPath = path;
            var source = VbaSourceHelper.ReadExportedTextAsUtf8(path);
            Component = new FakeComponent(2) { Name = "Widget", ImportedSource = source };
            Component.CodeModule.Text = VbaSourceHelper.NormalizeDocumentModuleContent(source);
            return Component;
        }

        public FakeComponent Add(int componentType)
        {
            AddCalls++;
            Component = new FakeComponent(componentType);
            return Component;
        }
    }

    private sealed class FakeComponent(int componentType)
    {
        public string Name { get; set; } = "";
        public int Type { get; } = componentType;
        public string ImportedSource { get; init; } = "";
        public ProvisioningCodeModule CodeModule { get; } = new();
        public int ExportCalls { get; private set; }

        public void Export(string path)
        {
            ExportCalls++;
            File.WriteAllText(path, ImportedSource, VbaSourceHelper.GetVbaImportEncoding());
        }
    }

    public sealed class ProvisioningCodeModule
    {
        public string Text { get; set; } = "";
        public int CountOfLines => string.IsNullOrEmpty(Text) ? 0 : Text.Split('\n').Length;
        public string Lines(object startLine, object count) => Text;

        public object? DeleteLines(object startLine, object count)
        {
            Text = "";
            return null;
        }

        public object? InsertLines(object startLine, object value)
        {
            Text = Convert.ToString(value) ?? "";
            return null;
        }
    }

    private sealed class FakeReferences
    {
        public List<(string LibID, int Major, int Minor)> Added { get; } = [];

        public object AddFromGuid(string libID, int major, int minor)
        {
            Added.Add((libID, major, minor));
            return new object();
        }
    }
}
