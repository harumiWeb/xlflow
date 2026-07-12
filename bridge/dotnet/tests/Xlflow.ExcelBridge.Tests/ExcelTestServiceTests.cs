using System.Text.Json;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelTestServiceTests
{
    [Fact]
    public void BuildErrorResponseIncludesWorkbookAndTestsInExtensions()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-error-ext",
            Command = "test",
        };
        var args = new TestCommandArguments(
            WorkbookPath: @"C:\work\book.xlsm",
            Filter: "",
            ModuleFilter: "",
            TagFilter: "",
            Visible: false,
            RuntimeMode: "",
            RuntimeSource: "",
            MsgBoxResponsesJSON: "",
            InputResponsesJSON: "",
            FileDialogResponsesJSON: "",
            DebugStreamEnabled: false,
            DebugStreamPipeName: "",
            UIStreamEnabled: false,
            UIStreamPipeName: "",
            UIStreamRedactInput: false,
            UseSession: true,
            MetadataPath: "");
        var tests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestSomething", Module = "Module1", Line = 5, Tags = ["smoke"] },
        };

        var response = ExcelTestService.BuildErrorResponse(
            request, args, "duplicate_test_name", "duplicate VBA test name(s): TestSomething",
            sessionAttached: true, sessionMode: "explicit", tests, runtimeState: null, runtimeInjected: false);

        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("duplicate_test_name", json.RootElement.GetProperty("error").GetProperty("code").GetString());
        Assert.Equal("test", json.RootElement.GetProperty("error").GetProperty("phase").GetString());
        Assert.Equal("xlflow", json.RootElement.GetProperty("error").GetProperty("source").GetString());

        var workbook = json.RootElement.GetProperty("workbook");
        Assert.Equal(@"C:\work\book.xlsm", workbook.GetProperty("path").GetString());
        Assert.True(workbook.GetProperty("session").GetBoolean());
        Assert.Equal("explicit", workbook.GetProperty("session_mode").GetString());
        Assert.True(workbook.GetProperty("session_requested").GetBoolean());
        Assert.False(workbook.GetProperty("auto_session").GetBoolean());

        var testsArray = json.RootElement.GetProperty("tests");
        Assert.Equal(1, testsArray.GetArrayLength());
        Assert.Equal("TestSomething", testsArray[0].GetProperty("name").GetString());
        Assert.Equal("Module1", testsArray[0].GetProperty("module").GetString());
        Assert.Equal("Module1.TestSomething", testsArray[0].GetProperty("id").GetString());
        Assert.Equal("Module1.TestSomething", testsArray[0].GetProperty("qualified_name").GetString());
        Assert.Equal(5, testsArray[0].GetProperty("line").GetInt32());

        var logs = json.RootElement.GetProperty("logs");
        Assert.Equal(1, logs.GetArrayLength());
        Assert.Contains("session", logs[0].GetString()!.ToLowerInvariant());
    }

    [Fact]
    public void BuildErrorResponseIncludesEmptyTestsForNoTestsFound()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-no-tests",
            Command = "test",
        };
        var args = new TestCommandArguments(
            WorkbookPath: @"C:\work\book.xlsm",
            Filter: "",
            ModuleFilter: "",
            TagFilter: "",
            Visible: false,
            RuntimeMode: "",
            RuntimeSource: "",
            MsgBoxResponsesJSON: "",
            InputResponsesJSON: "",
            FileDialogResponsesJSON: "",
            DebugStreamEnabled: false,
            DebugStreamPipeName: "",
            UIStreamEnabled: false,
            UIStreamPipeName: "",
            UIStreamRedactInput: false,
            UseSession: false,
            MetadataPath: "");

        var response = ExcelTestService.BuildErrorResponse(
            request, args, "no_tests_found", "no VBA tests found",
            sessionAttached: false, sessionMode: "none", [], runtimeState: null, runtimeInjected: false);

        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("no_tests_found", json.RootElement.GetProperty("error").GetProperty("code").GetString());

        var workbook = json.RootElement.GetProperty("workbook");
        Assert.Equal(@"C:\work\book.xlsm", workbook.GetProperty("path").GetString());
        Assert.False(workbook.GetProperty("session").GetBoolean());
        Assert.Equal("none", workbook.GetProperty("session_mode").GetString());

        var testsArray = json.RootElement.GetProperty("tests");
        Assert.Equal(0, testsArray.GetArrayLength());
    }

    [Fact]
    public void FindTestProceduresDiscoversTestPrefixAndSuffix()
    {
        const string code = """
            Option Explicit

            Public Sub TestAddition()
            End Sub

            Public Sub Helper()
            End Sub

            Public Sub Regression_Test()
            End Sub
            """;

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Equal(2, tests.Count);
        Assert.Equal("TestAddition", tests[0].Name);
        Assert.Equal("Module1", tests[0].Module);
        Assert.Equal("Module1.TestAddition", tests[0].Id);
        Assert.Equal("Module1.TestAddition", tests[0].QualifiedName);
        Assert.Equal("Regression_Test", tests[1].Name);
    }

    [Fact]
    public void FindTestProceduresDiscoversUnicodeTestPrefixAndSuffix()
    {
        const string code = """
            Option Explicit

            Public Sub Test_集計結果が正しい()
            End Sub

            Public Sub 集計結果_Test()
            End Sub

            Public Sub Helper_集計()
            End Sub
            """;

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Equal(2, tests.Count);
        Assert.Equal("Test_集計結果が正しい", tests[0].Name);
        Assert.Equal("集計結果_Test", tests[1].Name);
    }

    [Fact]
    public void FindTestProceduresCollectsTagAnnotations()
    {
        const string code = """
            Option Explicit

            ' @Tag("smoke")
            ' @Tag("fast")
            Public Sub TestWithTags()
            End Sub
            """;

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Single(tests);
        Assert.Equal("TestWithTags", tests[0].Name);
        Assert.Equal(2, tests[0].Tags.Length);
        Assert.Contains("smoke", tests[0].Tags);
        Assert.Contains("fast", tests[0].Tags);
    }

    [Fact]
    public void FindTestProceduresCollectsUnicodeTagAnnotations()
    {
        const string code = """
            Option Explicit

            ' @Tag("集計")
            Public Sub Test_集計結果が正しい()
            End Sub
            """;

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Single(tests);
        Assert.Equal("Test_集計結果が正しい", tests[0].Name);
        Assert.Equal(["集計"], tests[0].Tags);
    }

    [Fact]
    public void DuplicateTestQualifiedNamesAllowsSameNameInDifferentModules()
    {
        var allTests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestExport", Module = "InvoiceTests" },
            new() { Name = "TestExport", Module = "CustomerTests" },
            new() { Name = "testexport", Module = "InvoiceTests" },
        };

        var duplicates = ExcelTestService.DuplicateTestQualifiedNames(allTests);

        Assert.Equal(["InvoiceTests.TestExport"], duplicates);
    }

    [Fact]
    public void SelectTestsFiltersByQualifiedNameNameModuleAndTag()
    {
        var allTests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestA", Module = "Mod1", Tags = ["smoke"] },
            new() { Name = "TestB", Module = "Mod2", Tags = ["fast"] },
            new() { Name = "TestC", Module = "Mod1", Tags = ["slow"] },
        };

        var byName = ExcelTestService.SelectTests(allTests, "TestA", "", "");
        Assert.False(byName.Ambiguous);
        Assert.Single(byName.Tests);
        Assert.Equal("TestA", byName.Tests[0].Name);

        var byQualifiedName = ExcelTestService.SelectTests(allTests, "mod1.testa", "", "");
        Assert.False(byQualifiedName.Ambiguous);
        Assert.Single(byQualifiedName.Tests);
        Assert.Equal("Mod1.TestA", byQualifiedName.Tests[0].QualifiedName);

        var byModule = ExcelTestService.SelectTests(allTests, "", "Mod1", "");
        Assert.Equal(2, byModule.Tests.Count);

        var byTag = ExcelTestService.SelectTests(allTests, "", "", "fast");
        Assert.Single(byTag.Tests);
        Assert.Equal("TestB", byTag.Tests[0].Name);

        var byAll = ExcelTestService.SelectTests(allTests, "TestC", "Mod1", "slow");
        Assert.Single(byAll.Tests);
        Assert.Equal("TestC", byAll.Tests[0].Name);
    }

    [Fact]
    public void SelectTestsReportsAmbiguousUnqualifiedFilterAfterModuleAndTag()
    {
        var allTests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestExport", Module = "InvoiceTests", Tags = ["smoke"] },
            new() { Name = "TestExport", Module = "CustomerTests", Tags = ["smoke"] },
            new() { Name = "TestExport", Module = "DraftTests", Tags = ["slow"] },
        };

        var ambiguous = ExcelTestService.SelectTests(allTests, "testexport", "", "smoke");

        Assert.True(ambiguous.Ambiguous);
        Assert.Empty(ambiguous.Tests);
        Assert.Equal(["CustomerTests.TestExport", "InvoiceTests.TestExport"], ambiguous.Matches.Select(t => t.QualifiedName).ToArray());

        var narrowed = ExcelTestService.SelectTests(allTests, "testexport", "invoiceTests", "smoke");
        Assert.False(narrowed.Ambiguous);
        Assert.Single(narrowed.Tests);
        Assert.Equal("InvoiceTests.TestExport", narrowed.Tests[0].QualifiedName);
    }

    [Fact]
    public void BuildErrorResponseIncludesAmbiguousMatchesInErrorDetails()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-ambiguous",
            Command = "test",
        };
        var args = new TestCommandArguments(
            WorkbookPath: @"C:\work\book.xlsm",
            Filter: "TestExport",
            ModuleFilter: "",
            TagFilter: "",
            Visible: false,
            RuntimeMode: "",
            RuntimeSource: "",
            MsgBoxResponsesJSON: "",
            InputResponsesJSON: "",
            FileDialogResponsesJSON: "",
            DebugStreamEnabled: false,
            DebugStreamPipeName: "",
            UIStreamEnabled: false,
            UIStreamPipeName: "",
            UIStreamRedactInput: false,
            UseSession: false,
            MetadataPath: "");
        var tests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestExport", Module = "InvoiceTests" },
            new() { Name = "TestExport", Module = "CustomerTests" },
        };

        var response = ExcelTestService.BuildErrorResponse(
            request,
            args,
            "ambiguous_test_name",
            "test name is ambiguous: TestExport",
            sessionAttached: false,
            sessionMode: "none",
            tests,
            runtimeState: null,
            runtimeInjected: false,
            errorDetails: new Dictionary<string, object?> { ["matches"] = tests.Select(t => t.QualifiedName).ToArray() });

        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ambiguous_test_name", json.RootElement.GetProperty("error").GetProperty("code").GetString());
        var matches = json.RootElement.GetProperty("error").GetProperty("details").GetProperty("matches");
        Assert.Equal(2, matches.GetArrayLength());
        Assert.Equal("InvoiceTests.TestExport", matches[0].GetString());
    }

    [Fact]
    public void AdjustCountsForAfterAllFailureConvertsInconclusiveToFailed()
    {
        var failed = 0;
        var inconclusive = 1;

        ExcelTestService.AdjustCountsForAfterAllFailure("inconclusive", ref failed, ref inconclusive);

        Assert.Equal(1, failed);
        Assert.Equal(0, inconclusive);
    }
}
