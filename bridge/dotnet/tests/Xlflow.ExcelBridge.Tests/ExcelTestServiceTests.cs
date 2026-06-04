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
        Assert.Equal("Regression_Test", tests[1].Name);
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
    public void SelectTestsFiltersByNameModuleAndTag()
    {
        var allTests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestA", Module = "Mod1", Tags = ["smoke"] },
            new() { Name = "TestB", Module = "Mod2", Tags = ["fast"] },
            new() { Name = "TestC", Module = "Mod1", Tags = ["slow"] },
        };

        var byName = ExcelTestService.SelectTests(allTests, "TestA", "", "");
        Assert.Single(byName);
        Assert.Equal("TestA", byName[0].Name);

        var byModule = ExcelTestService.SelectTests(allTests, "", "Mod1", "");
        Assert.Equal(2, byModule.Count);

        var byTag = ExcelTestService.SelectTests(allTests, "", "", "fast");
        Assert.Single(byTag);
        Assert.Equal("TestB", byTag[0].Name);

        var byAll = ExcelTestService.SelectTests(allTests, "TestC", "Mod1", "slow");
        Assert.Single(byAll);
        Assert.Equal("TestC", byAll[0].Name);
    }
}
