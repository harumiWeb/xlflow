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
