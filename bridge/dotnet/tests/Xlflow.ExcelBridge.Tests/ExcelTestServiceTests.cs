using System.Runtime.InteropServices;
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
            new() { Name = "TestSomething", Module = "Module1", Line = 5, Tags = ["smoke"], ExpectedError = new ExcelTestService.ExpectedErrorMetadata { Number = 5 } },
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
        Assert.Equal(5, testsArray[0].GetProperty("expected_error").GetProperty("number").GetInt32());

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
    public void BuildErrorResponseIncludesSkipTodoDiscoveryMetadata()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-status-metadata",
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
        var tests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestAccess", Module = "Module1", Skip = new ExcelTestService.StatusReasonMetadata { Reason = "Requires Access" } },
            new() { Name = "TestExporter", Module = "Module1", Todo = new ExcelTestService.StatusReasonMetadata() },
        };

        var response = ExcelTestService.BuildErrorResponse(
            request, args, "test_not_found", "test not found",
            sessionAttached: false, sessionMode: "none", tests, runtimeState: null, runtimeInjected: false);

        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);
        var testsArray = json.RootElement.GetProperty("tests");

        Assert.Equal("skipped", testsArray[0].GetProperty("status_hint").GetString());
        Assert.Equal("Requires Access", testsArray[0].GetProperty("skip").GetProperty("reason").GetString());
        Assert.Equal("todo", testsArray[1].GetProperty("status_hint").GetString());
        Assert.Equal(JsonValueKind.Object, testsArray[1].GetProperty("todo").ValueKind);
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
        const string code = "Option Explicit\r\n\r\n" +
            "' @Tag(\"smoke\")\r\n" +
            "' @ExpectedError(5, \"Invalid \"\"value\"\"\", \"ParserModule\")\r\n" +
            "' @Tag(\"fast\")\r\n" +
            "Public Sub TestWithTags()\r\n" +
            "End Sub\r\n";

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Single(tests);
        Assert.Equal("TestWithTags", tests[0].Name);
        Assert.Equal(2, tests[0].Tags.Length);
        Assert.Contains("smoke", tests[0].Tags);
        Assert.Contains("fast", tests[0].Tags);
        var expectedError = tests[0].ExpectedError;
        Assert.NotNull(expectedError);
        Assert.Equal(5, expectedError.Number);
        Assert.Equal("Invalid \"value\"", expectedError.Description);
        Assert.Equal("ParserModule", expectedError.Source);
    }

    [Fact]
    public void FindTestProceduresExpandsParameterizedTestCases()
    {
        const string code = "Option Explicit\r\n\r\n" +
            "' @Tag(\"fast\")\r\n" +
            "' @TestCase(\"positive\"; 1, 2.5, True, \"a \"\"quote\"\"\", Empty, Null, #2026-07-12#)\r\n" +
            "' @TestCase(-1, 1.2E-3, False, \"\", Empty, Null, #2026-07-12#)\r\n" +
            "Public Sub Test_Add( _\r\n" +
            "    ByVal leftValue As Long, _\r\n" +
            "    ByVal rightValue As Double, _\r\n" +
            "    ByVal enabled As Boolean, _\r\n" +
            "    ByVal label As String, _\r\n" +
            "    ByVal emptyValue As Variant, _\r\n" +
            "    ByVal nullValue As Variant, _\r\n" +
            "    ByVal dayValue As Date)\r\n" +
            "End Sub\r\n";

        var tests = ExcelTestService.FindTestProcedures("MathTests", code);

        Assert.Equal(2, tests.Count);
        Assert.Equal("MathTests.Test_Add[positive]", tests[0].QualifiedName);
        Assert.Equal("MathTests.Test_Add", tests[0].QualifiedProcedure);
        Assert.Equal("positive", tests[0].CaseId);
        Assert.Equal(4, tests[0].AnnotationLine);
        Assert.Equal(6, tests[0].ProcedureLine);
        Assert.Equal(7, tests[0].Arguments.Length);
        Assert.Equal("Long", tests[0].Arguments[0].Type);
        Assert.Equal(1L, tests[0].Arguments[0].Value);
        Assert.Equal("a \"quote\"", tests[0].Arguments[3].Value);
        Assert.Null(tests[0].Arguments[5].Value);
        Assert.Equal("2026-07-12", tests[0].Arguments[6].Value);
        Assert.Equal("MathTests.Test_Add[-1,1.2E-3,False,\"\",Empty,Null,#2026-07-12#]", tests[1].QualifiedName);
    }

    [Theory]
    [InlineData("' @TestCase(1)\r\nPublic Sub Test_Add(ByVal leftValue As Long, ByVal rightValue As Long)\r\nEnd Sub", "provides 1 arguments")]
    [InlineData("' @TestCase(\"abc\")\r\nPublic Sub Test_Parse(ByVal value As Long)\r\nEnd Sub", "cannot be passed to Long")]
    [InlineData("' @TestCase(\"A1\")\r\nPublic Sub Test_Range(ByVal value As Range)\r\nEnd Sub", "unsupported parameter type Range")]
    [InlineData("' @TestCase(1)\r\nPublic Sub Test_ByRef(ByRef value As Long)\r\nEnd Sub", "must be ByVal")]
    [InlineData("' @TestCase(SomeConstant)\r\nPublic Sub Test_Constant(ByVal value As Long)\r\nEnd Sub", "unsupported @TestCase literal")]
    public void FindTestProceduresRejectsInvalidTestCases(string body, string expectedMessage)
    {
        var code = "Option Explicit\r\n\r\n" + body;

        var ex = Assert.Throws<ExcelTestService.InvalidTestCaseException>(() =>
            ExcelTestService.FindTestProcedures("Module1", code));

        Assert.Contains(expectedMessage, ex.Message);
    }

    [Fact]
    public void FindTestProceduresRejectsDuplicateGeneratedCaseIds()
    {
        const string code = """
            Option Explicit

            ' @TestCase("same"; 1)
            ' @TestCase("same"; 2)
            Public Sub Test_Dupe(ByVal value As Long)
            End Sub
            """;

        var ex = Assert.Throws<ExcelTestService.InvalidTestCaseException>(() =>
            ExcelTestService.FindTestProcedures("Module1", code));

        Assert.Contains("duplicate generated test case id", ex.Message);
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
    public void FindTestProceduresCollectsSkipAndTodoAnnotations()
    {
        const string code = """
            Option Explicit

            ' @Skip("Requires Access")
            Public Sub TestAccess()
            End Sub

            ' @Todo
            Public Sub TestExporter()
            End Sub

            ' @Todo("実装待ち")
            Public Sub TestUnicode()
            End Sub
            """;

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Equal(3, tests.Count);
        Assert.NotNull(tests[0].Skip);
        Assert.Equal("Requires Access", tests[0].Skip!.Reason);
        Assert.Null(tests[0].Todo);
        Assert.NotNull(tests[1].Todo);
        Assert.Null(tests[1].Todo!.Reason);
        Assert.Equal("実装待ち", tests[2].Todo!.Reason);
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
    public void FindTestProceduresCollectsExpectedErrorNumberOnly()
    {
        const string code = """
            Option Explicit

            ' @ExpectedError(5)
            Public Sub TestInvalidArgument()
            End Sub
            """;

        var tests = ExcelTestService.FindTestProcedures("Module1", code);

        Assert.Single(tests);
        var expectedError = tests[0].ExpectedError;
        Assert.NotNull(expectedError);
        Assert.Equal(5, expectedError.Number);
        Assert.Null(expectedError.Description);
        Assert.Null(expectedError.Source);
    }

    [Theory]
    [InlineData("' @ExpectedError(foo)", "@ExpectedError error number must be numeric")]
    [InlineData("' @ExpectedError(5, \"a\", \"b\", \"c\")", "@ExpectedError supports 1 to 3 arguments")]
    [InlineData("' @ExpectedError(5, \"unterminated)", "malformed string literal")]
    public void FindTestProceduresRejectsMalformedExpectedError(string annotation, string expectedMessage)
    {
        var code = $$"""
            Option Explicit

            {{annotation}}
            Public Sub TestInvalidArgument()
            End Sub
            """;

        var ex = Assert.Throws<ExcelTestService.InvalidTestMetadataException>(() =>
            ExcelTestService.FindTestProcedures("Module1", code));

        Assert.Contains(expectedMessage, ex.Message);
        Assert.Contains("Module1:3", ex.Message);
    }

    [Fact]
    public void FindTestProceduresRejectsDuplicateExpectedError()
    {
        const string code = """
            Option Explicit

            ' @ExpectedError(5)
            ' @ExpectedError(6)
            Public Sub TestInvalidArgument()
            End Sub
            """;

        var ex = Assert.Throws<ExcelTestService.InvalidTestMetadataException>(() =>
            ExcelTestService.FindTestProcedures("Module1", code));

        Assert.Contains("multiple @ExpectedError", ex.Message);
        Assert.Contains("Module1:3", ex.Message);
    }

    [Theory]
    [InlineData("' @Skip(\"a\")\r\n' @Skip(\"b\")", "multiple @Skip")]
    [InlineData("' @Todo(\"a\")\r\n' @Todo(\"b\")", "multiple @Todo")]
    [InlineData("' @Skip(\"a\")\r\n' @Todo(\"b\")", "test cannot be both skipped and todo")]
    [InlineData("' @Skip()", "malformed @Skip reason")]
    public void FindTestProceduresRejectsInvalidSkipTodoMetadata(string annotations, string expectedMessage)
    {
        var code = $$"""
            Option Explicit

            {{annotations}}
            Public Sub TestInvalidMetadata()
            End Sub
            """;

        var ex = Assert.Throws<ExcelTestService.InvalidTestMetadataException>(() =>
            ExcelTestService.FindTestProcedures("Module1", code));

        Assert.Contains(expectedMessage, ex.Message);
        Assert.Contains("Module1:", ex.Message);
    }

    [Fact]
    public void FindTestProceduresRejectsExpectedErrorOnNonTestProcedure()
    {
        const string code = """
            Option Explicit

            ' @ExpectedError(5)
            Public Sub Helper()
            End Sub
            """;

        var ex = Assert.Throws<ExcelTestService.InvalidTestMetadataException>(() =>
            ExcelTestService.FindTestProcedures("Module1", code));

        Assert.Contains("only supported on test procedures", ex.Message);
        Assert.Contains("Module1:3", ex.Message);
    }

    [Fact]
    public void ClassifyTestOutcomeMatchesExpectedError()
    {
        var test = new ExcelTestService.TestCase
        {
            Name = "TestInvalidArgument",
            Module = "ParserTests",
            ExpectedError = new ExcelTestService.ExpectedErrorMetadata { Number = 5 },
        };

        var result = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: 5, errSource: "ParserModule", errDescription: "Invalid value", statusHint: "", phaseHint: "test");

        Assert.Equal("passed", result.Status);
        Assert.NotNull(result.ObservedError);
    }

    [Fact]
    public void ClassifyTestOutcomeFailsWhenExpectedErrorIsMissing()
    {
        var test = new ExcelTestService.TestCase
        {
            Name = "TestInvalidArgument",
            Module = "ParserTests",
            ExpectedError = new ExcelTestService.ExpectedErrorMetadata { Number = 5 },
        };

        var result = ExcelTestService.ClassifyTestOutcome(test, success: true, errNumber: 0, errSource: "", errDescription: "", statusHint: "", phaseHint: "");

        Assert.Equal("failed", result.Status);
        Assert.Equal("expected_error_mismatch", result.ErrorCode);
        Assert.Equal("expected VBA error 5 but no error was raised", result.ErrorMessage);
    }

    [Fact]
    public void ClassifyTestOutcomeFailsExpectedErrorNumberDescriptionAndSourceMismatches()
    {
        var test = new ExcelTestService.TestCase
        {
            Name = "TestInvalidArgument",
            Module = "ParserTests",
            ExpectedError = new ExcelTestService.ExpectedErrorMetadata { Number = 5, Description = "Invalid value", Source = "InvoiceParser" },
        };

        var wrongNumber = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: 9, errSource: "InvoiceParser", errDescription: "Subscript out of range", statusHint: "", phaseHint: "test");
        Assert.Equal("expected_error_mismatch", wrongNumber.ErrorCode);
        Assert.Contains("expected VBA error 5 but got error 9", wrongNumber.ErrorMessage);

        var wrongDescription = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: 5, errSource: "InvoiceParser", errDescription: "Value is invalid", statusHint: "", phaseHint: "test");
        Assert.Equal("expected error description <Invalid value> but got <Value is invalid>", wrongDescription.ErrorMessage);

        var wrongSource = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: 5, errSource: "ParserModule", errDescription: "Invalid value", statusHint: "", phaseHint: "test");
        Assert.Equal("expected error source <InvoiceParser> but got <ParserModule>", wrongSource.ErrorMessage);
    }

    [Fact]
    public void ClassifyTestOutcomeMatchesSourceCaseInsensitively()
    {
        var test = new ExcelTestService.TestCase
        {
            Name = "TestInvalidArgument",
            Module = "ParserTests",
            ExpectedError = new ExcelTestService.ExpectedErrorMetadata { Number = 5, Source = "InvoiceParser" },
        };

        var result = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: 5, errSource: "invoiceparser", errDescription: "Invalid value", statusHint: "", phaseHint: "test");

        Assert.Equal("passed", result.Status);
    }

    [Fact]
    public void ClassifyTestOutcomeKeepsHookAndInconclusivePrecedence()
    {
        var test = new ExcelTestService.TestCase
        {
            Name = "TestInvalidArgument",
            Module = "ParserTests",
            ExpectedError = new ExcelTestService.ExpectedErrorMetadata { Number = unchecked((int)0x80040000) + 516 },
        };

        var inconclusive = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: unchecked((int)0x80040000) + 516, errSource: "XlflowAssert.AssertInconclusive", errDescription: "not ready", statusHint: "inconclusive", phaseHint: "test");
        Assert.Equal("inconclusive", inconclusive.Status);
        Assert.Equal("test_inconclusive", inconclusive.ErrorCode);

        var beforeEach = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: unchecked((int)0x80040000) + 516, errSource: "BeforeEach", errDescription: "setup failed", statusHint: "", phaseHint: "before_each");
        Assert.Equal("failed", beforeEach.Status);
        Assert.Equal("before_each_failed", beforeEach.ErrorCode);

        var afterEach = ExcelTestService.ClassifyTestOutcome(test, success: false, errNumber: unchecked((int)0x80040000) + 516, errSource: "AfterEach", errDescription: "cleanup failed", statusHint: "failed", phaseHint: "after_each");
        Assert.Equal("failed", afterEach.Status);
        Assert.Equal("after_each_failed", afterEach.ErrorCode);
    }

    [Fact]
    public void CountFailedResultsReadsDictionaryTestPayloads()
    {
        var response = new BridgeResponse
        {
            Status = BridgeStatus.Failed,
            Error = new BridgeError("test_failed", "1 failed", "test", "xlflow"),
            Extensions = new Dictionary<string, object?>
            {
                ["tests"] = new object[]
                {
                    new Dictionary<string, object?> { ["status"] = "passed" },
                    new Dictionary<string, object?> { ["status"] = "failed" },
                    new Dictionary<string, object?> { ["status"] = "skipped" },
                    new Dictionary<string, object?> { ["status"] = "todo" },
                    new Dictionary<string, object?> { ["status"] = "inconclusive" },
                },
            },
        };

        Assert.Equal(1, ExcelTestService.CountFailedResults(response));
    }

    [Fact]
    public void CountFailedResultsFallsBackWhenTestsPayloadIsMissing()
    {
        var response = new BridgeResponse
        {
            Status = BridgeStatus.Failed,
            Error = new BridgeError("test_failed", "test failed", "test", "xlflow"),
        };

        Assert.Equal(1, ExcelTestService.CountFailedResults(response));
    }

    [Fact]
    public void BuildTestResultOmitsErrorForPassedTests()
    {
        var test = new ExcelTestService.TestCase { Name = "TestOk", Module = "ParserTests" };

        var passed = ExcelTestService.BuildTestResult(test, "passed", 5,
            new { code = "", message = "", source = "", number = 0 });
        var failed = ExcelTestService.BuildTestResult(test, "failed", 5,
            new { code = "test_failed", message = "failed", source = "ParserTests", number = 5 });

        Assert.False(passed.ContainsKey("error"));
        Assert.True(failed.ContainsKey("error"));
        Assert.Equal(1, passed["attempts"]);
        Assert.Equal(false, passed["flaky"]);
        Assert.True(passed.ContainsKey("attempt_results"));
    }

    [Fact]
    public void BuildNotRunTestResultOmitsErrorAndIncludesReason()
    {
        var test = new ExcelTestService.TestCase
        {
            Name = "TestLater",
            Module = "ParserTests",
            CaseId = "case-a",
            AnnotationLine = 12,
        };

        var result = ExcelTestService.BuildNotRunTestResult(test, "maximum failure count reached");

        Assert.Equal("ParserTests.TestLater[case-a]", result["qualified_name"]);
        Assert.Equal("not_run", result["status"]);
        Assert.Equal("maximum failure count reached", result["reason"]);
        Assert.Equal(0, result["duration_ms"]);
        Assert.False(result.ContainsKey("error"));
        Assert.Equal(1, result["attempts"]);
        Assert.Equal(false, result["flaky"]);
        var attempts = Assert.IsAssignableFrom<IEnumerable<object>>(result["attempt_results"]);
        var attempt = Assert.Single(attempts);
        var attemptPayload = Assert.IsAssignableFrom<IReadOnlyDictionary<string, object?>>(attempt);
        Assert.Equal("not_run", attemptPayload["status"]);
        Assert.Equal(0, attemptPayload["duration_ms"]);
    }

    [Fact]
    public void BuildNonExecutedTestResultOmitsErrorAndIncludesReason()
    {
        var skippedTest = new ExcelTestService.TestCase
        {
            Name = "TestAccess",
            Module = "ParserTests",
            Skip = new ExcelTestService.StatusReasonMetadata { Reason = "Requires Access" },
        };
        var todoTest = new ExcelTestService.TestCase
        {
            Name = "TestExporter",
            Module = "ParserTests",
            Todo = new ExcelTestService.StatusReasonMetadata(),
        };

        var skipped = ExcelTestService.BuildNonExecutedTestResult(skippedTest);
        var todo = ExcelTestService.BuildNonExecutedTestResult(todoTest);

        Assert.Equal("skipped", skipped["status"]);
        Assert.Equal("Requires Access", skipped["reason"]);
        Assert.False(skipped.ContainsKey("error"));
        Assert.Equal("todo", todo["status"]);
        Assert.False(todo.ContainsKey("reason"));
        Assert.False(todo.ContainsKey("error"));
    }

    [Fact]
    public void SelectTestsFiltersByQualifiedNameNameModuleAndTag()
    {
        var allTests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestA", Module = "Mod1", Tags = ["smoke"] },
            new() { Name = "TestB", Module = "Mod2", Tags = ["fast"], Todo = new ExcelTestService.StatusReasonMetadata { Reason = "pending" } },
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
        Assert.NotNull(byTag.Tests[0].Todo);

        var byAll = ExcelTestService.SelectTests(allTests, "TestC", "Mod1", "slow");
        Assert.Single(byAll.Tests);
        Assert.Equal("TestC", byAll.Tests[0].Name);
    }

    [Fact]
    public void SelectTestsFiltersParameterizedCasesByProcedureOrExactCase()
    {
        var allTests = new List<ExcelTestService.TestCase>
        {
            new() { Name = "TestAdd", Module = "MathTests", CaseId = "1,2,3" },
            new() { Name = "TestAdd", Module = "MathTests", CaseId = "-1,1,0" },
            new() { Name = "TestOther", Module = "MathTests" },
        };

        var procedure = ExcelTestService.SelectTests(allTests, "MathTests.TestAdd", "", "");
        Assert.Equal(2, procedure.Tests.Count);

        var exactCase = ExcelTestService.SelectTests(allTests, "MathTests.TestAdd[1,2,3]", "", "");
        Assert.Single(exactCase.Tests);
        Assert.Equal("1,2,3", exactCase.Tests[0].CaseId);
    }

    [Fact]
    public void BuildTestRunnerCodePassesParameterizedArguments()
    {
        var tests = new List<ExcelTestService.TestCase>
        {
            new()
            {
                Name = "TestAdd",
                Module = "MathTests",
                CaseId = "positive",
                Arguments =
                [
                    new ExcelTestService.TestArgument { Type = "Long", Value = 1L, VbaLiteral = "1" },
                    new ExcelTestService.TestArgument { Type = "String", Value = "a \"quote\"", VbaLiteral = "\"a \"\"quote\"\"\"" },
                ],
            },
        };

        var code = ExcelTestService.BuildTestRunnerCode(tests, []);

        Assert.Contains("MathTests.TestAdd 1, \"a \"\"quote\"\"\"", code);
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

    [Theory]
    [InlineData("skipped")]
    [InlineData("todo")]
    [InlineData("not_run")]
    public void AdjustCountsForAfterAllFailureIgnoresNonExecutedStatuses(string status)
    {
        var failed = 0;
        var inconclusive = 0;

        ExcelTestService.AdjustCountsForAfterAllFailure(status, ref failed, ref inconclusive);

        Assert.Equal(0, failed);
        Assert.Equal(0, inconclusive);
    }

    [Fact]
    public void CopyWorkbookForTestPreservesExtensionAndUsesUniqueNames()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var source = Path.Combine(dir, "Book.xlsm");
            File.WriteAllText(source, "test workbook");
            var runDir = Path.Combine(dir, ".xlflow", "test-runs", "abc123");

            var first = ExcelTestService.CopyWorkbookForTest(source, runDir, "Module.Tests");
            var second = ExcelTestService.CopyWorkbookForTest(source, runDir, "Module.Tests");

            Assert.EndsWith(".xlsm", first);
            Assert.EndsWith(".xlsm", second);
            Assert.NotEqual(first, second);
            Assert.StartsWith(runDir, first, StringComparison.OrdinalIgnoreCase);
            Assert.True(File.Exists(first));
            Assert.True(File.Exists(second));
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
    }

    [Fact]
    public void FatalComFailureProducesTypedRecoveryOutcome()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-fatal-com",
            Command = "test",
        };
        var exception = new COMException("RPC disconnected", unchecked((int)0x80010108));

        Assert.True(ExcelTestService.IsFatalTestComFailure(exception));

        var response = ExcelTestService.BuildUnhandledTestFailureResponse(
            request,
            exception,
            excelProcessId: 24680,
            sessionAttached: true,
            sessionMode: "managed");

        Assert.Equal(BridgeStatus.Failed, response.Status);
        Assert.Equal("excel_com_rpc_failure", response.Error?.Code);
        Assert.Equal("0x80010108", response.Error?.HResult);
        Assert.NotNull(response.Recovery);
        Assert.True(response.Recovery!.Required);
        Assert.Equal("excel_com_state_uncertain", response.Recovery.Reason);
        Assert.Equal("test", response.Recovery.Operation);
        Assert.Equal(24680, response.Recovery.ExcelProcessId);
        Assert.True(response.Recovery.Session.Active);
        Assert.Equal("managed", response.Recovery.Session.Owner);
    }

    [Fact]
    public void OrdinaryTestEnvironmentFailureDoesNotProduceRecovery()
    {
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-test-ordinary-failure",
            Command = "test",
        };
        var exception = new InvalidOperationException("ordinary test error");

        Assert.False(ExcelTestService.IsFatalTestComFailure(exception));

        var response = ExcelTestService.BuildUnhandledTestFailureResponse(
            request,
            exception,
            excelProcessId: 0,
            sessionAttached: false,
            sessionMode: "none");

        Assert.Equal("test_environment_failed", response.Error?.Code);
        Assert.Null(response.Recovery);
    }
}
