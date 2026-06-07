using System.Globalization;
using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;
using Xlflow.ExcelBridge.Windows;
using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Tests;

public sealed class RunCommandTests
{
    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new RunCommand(new FakeRunService((_, _) => BridgeResponse.Ok(new BridgeRequest())));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-run-missing-path",
            Command = "run",
            Payload = JsonDocument.Parse("""
                {
                  "MacroName": "Module1.Main"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("run_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandleRejectsMissingMacroName()
    {
        var command = new RunCommand(new FakeRunService((_, _) => BridgeResponse.Ok(new BridgeRequest())));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-run-missing-macro",
            Command = "run",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("run_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandlePassesUiOrchestrationOptionsToService()
    {
        var serviceCalled = false;
        var command = new RunCommand(new FakeRunService((request, args) =>
        {
            serviceCalled = true;
            Assert.Equal("[]", args.MsgBoxResponsesJSON);
            Assert.Equal("{\"customer-name\":\"Jane\"}", args.InputResponsesJSON);
            Assert.Equal("[{\"kind\":\"file\",\"dialog_id\":\"pick-report\"}]", args.FileDialogResponsesJSON);
            Assert.True(args.UIStreamEnabled);
            Assert.True(args.UIStreamRedactInput);
            Assert.True(args.DebugStreamEnabled);
            Assert.Equal(@"\\.\pipe\xlflow-debug-test", args.DebugStreamPipeName);
            Assert.Equal(@"C:\work\src\modules", args.ModulesDir);
            Assert.Equal(@"C:\work\src\classes", args.ClassesDir);
            Assert.Equal(@"C:\work\src\forms", args.FormsDir);
            Assert.Equal(@"C:\work\src\workbook", args.WorkbookDir);
            Assert.Equal("sidecar", args.CodeSource);
            Assert.True(args.Folders);
            Assert.Equal("update", args.FolderAnnotation);
            Assert.True(args.DefaultComponentFolders);

            return BridgeResponse.Ok(request);
        }));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-run-ui-supported",
            Command = "run",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "MacroName": "Module1.Main",
                  "MsgBoxResponsesJSON": "[]",
                  "InputResponsesJSON": "{\"customer-name\":\"Jane\"}",
                  "FileDialogResponsesJSON": "[{\"kind\":\"file\",\"dialog_id\":\"pick-report\"}]",
                  "UIStreamEnabled": true,
                  "UIStreamRedactInput": true,
                  "DebugStreamEnabled": true,
                  "DebugStreamPipeName": "\\\\.\\pipe\\xlflow-debug-test",
                  "ModulesDir": "C:\\work\\src\\modules",
                  "ClassesDir": "C:\\work\\src\\classes",
                  "FormsDir": "C:\\work\\src\\forms",
                  "WorkbookDir": "C:\\work\\src\\workbook",
                  "CodeSource": "sidecar",
                  "Folders": true,
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": true
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.True(serviceCalled);
        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void InvokeMethodUsesCurrentCultureForComLateBinding()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        var originalUiCulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("fr-FR");

            Assert.Equal("fr-FR", ExcelBridgeSupport.ComInvokeCulture.Name);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public void HandleRejectsSaveAsWithMismatchedExtension()
    {
        var service = new FakeRunService((request, args) =>
        {
            try
            {
                ExcelRunService.AssertSaveAsExtension(args.WorkbookPath, args.SaveAsPath);
                return BridgeResponse.Ok(request);
            }
            catch (InvalidOperationException ex)
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "save_as_extension_mismatch",
                    Message: ex.Message,
                    Phase: "save_as",
                    Source: "xlflow-excel-bridge"));
            }
        });
        var command = new RunCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-run-saveas-ext",
            Command = "run",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "MacroName": "Module1.Main",
                  "SaveAsPath": "C:\\work\\book.xlsx"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        var error = json.RootElement.GetProperty("error");
        Assert.Contains("extension", error.GetProperty("message").GetString()!.ToLowerInvariant());
    }

    [Fact]
    public void HandleAcceptsSaveAsWithMatchingExtension()
    {
        var service = new FakeRunService((request, args) =>
        {
            ExcelRunService.AssertSaveAsExtension(args.WorkbookPath, args.SaveAsPath);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["saved"] = true,
                    ["save_as"] = args.SaveAsPath,
                },
            });
        });
        var command = new RunCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-run-saveas-ok",
            Command = "run",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "MacroName": "Module1.Main",
                  "SaveAsPath": "C:\\work\\book_copy.xlsm"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroNotFoundForCannotRunMacro()
    {
        Assert.Equal("macro_not_found", ExcelRunService.ClassifyRunError("Cannot run the macro", null));
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroNotFoundForSubOrFunctionNotDefined()
    {
        Assert.Equal("macro_not_found", ExcelRunService.ClassifyRunError("Sub or function not defined", null));
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroNotFoundFor1004WithMacro()
    {
        Assert.Equal("macro_not_found", ExcelRunService.ClassifyRunError("some macro error", 1004));
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroDisabledForSecuritySettings()
    {
        Assert.Equal("macro_disabled", ExcelRunService.ClassifyRunError("Security settings prevent macro", null));
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroDisabledForMacroDisabled()
    {
        Assert.Equal("macro_disabled", ExcelRunService.ClassifyRunError("Macro is disabled", null));
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroFailedForUnrecognizedError()
    {
        Assert.Equal("macro_failed", ExcelRunService.ClassifyRunError("Something went wrong", null));
    }

    [Fact]
    public void ClassifyRunErrorReturnsMacroFailedForNullMessage()
    {
        Assert.Equal("macro_failed", ExcelRunService.ClassifyRunError("", null));
    }

    [Fact]
    public void FatalComFailureClassifierDetectsRpcFailureAndFormatsHResult()
    {
        const int rpcFailed = unchecked((int)0x800706BE);

        Assert.True(ExcelBridgeSupport.IsFatalComFailure(rpcFailed));
        Assert.Equal("0x800706BE", ExcelBridgeSupport.FormatHResult(rpcFailed));
    }

    [Fact]
    public void BuildRunHarnessCodePumpsDoEventsAroundMacroInvocation()
    {
        var code = ExcelRunService.BuildRunHarnessCode("Main.ReproFormatOps", []);

        Assert.Contains("  DoEvents" + Environment.NewLine + "  Application.Run targetMacro", code);
        Assert.Contains("  Application.Run targetMacro" + Environment.NewLine + "  DoEvents", code);
    }

    [Fact]
    public void ComFailureDetailsIncludeRunContext()
    {
        var args = RunArgs("C:\\work\\Book.xlsm", "Main.ReproFormatOps");

        var details = ExcelRunService.BuildComFailureDetails(
            args,
            "application_run",
            excelProcessId: 1234,
            excelHwnd: 5678,
            workerProcessId: 9012,
            sessionAttached: true,
            sessionMode: "explicit");

        Assert.Equal("Main.ReproFormatOps", details["macro"]);
        Assert.Equal("application_run", details["stage"]);
        Assert.Equal(1234, details["excel_pid"]);
        Assert.Equal(5678L, details["excel_hwnd"]);
        Assert.Equal(9012, details["worker_pid"]);
        Assert.Equal("explicit", details["session_mode"]);
        Assert.Equal(false, details["visible"]);
        Assert.Equal(true, details["headless"]);
        Assert.Equal(true, details["workbook_reused"]);
        Assert.Equal("explicit", details["workbook_reuse_mode"]);
    }

    [Theory]
    [InlineData(true, false, true, true)]
    [InlineData(false, false, false, false)]
    [InlineData(false, true, false, false)]
    public void ComputePostRunSaveStateMatchesSessionPersistenceRules(bool sessionAttached, bool saved, bool expectedDirty, bool expectedNeedsSave)
    {
        var result = ExcelRunService.ComputePostRunSaveState(sessionAttached, saved);

        Assert.Equal(expectedDirty, result.Dirty);
        Assert.Equal(expectedNeedsSave, result.NeedsSave);
    }

    [Fact]
    public void DefaultRuntimeSkipRequiresNoDebugOrUiInjection()
    {
        var args = RunArgs(@"C:\work\Book.xlsm", "Main.Run") with
        {
            RuntimeSource = "default",
        };

        Assert.True(ExcelRunService.IsDefaultRuntimeWithoutInjectedFeatures(args));

        args = args with
        {
            DebugStreamEnabled = true,
            DebugStreamPipeName = @"\\.\pipe\xlflow-debug",
        };
        Assert.False(ExcelRunService.IsDefaultRuntimeWithoutInjectedFeatures(args));

        args = args with
        {
            DebugStreamEnabled = false,
            DebugStreamPipeName = "",
            UIStreamEnabled = true,
            UIStreamPipeName = @"\\.\pipe\xlflow-ui",
        };
        Assert.False(ExcelRunService.IsDefaultRuntimeWithoutInjectedFeatures(args));
    }

    private static RunCommandArguments RunArgs(string workbookPath, string macroName)
    {
        return new RunCommandArguments(
            WorkbookPath: workbookPath,
            MacroName: macroName,
            MacroArgsJSON: "",
            Visible: false,
            DisplayAlerts: false,
            SaveWorkbook: false,
            Direct: false,
            Diagnostic: false,
            SuppressModalErrors: true,
            MsgBoxResponsesJSON: "",
            InputResponsesJSON: "",
            FileDialogResponsesJSON: "",
            DebugStreamEnabled: false,
            DebugStreamPipeName: "",
            UIStreamEnabled: false,
            UIStreamPipeName: "",
            UIStreamRedactInput: false,
            UseSession: true,
            MetadataPath: "C:\\work\\.xlflow\\session.json",
            RuntimeMode: "",
            RuntimeSource: "",
            SaveAsPath: "",
            TimeoutSeconds: 0,
            ModulesDir: "",
            ClassesDir: "",
            FormsDir: "",
            WorkbookDir: "",
            CodeSource: "",
            Folders: false,
            FolderAnnotation: "update",
            DefaultComponentFolders: false);
    }

    [Fact]
    public void Execute_ReturnsBridgeFileNotOpenableForNonExcelFile()
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "xlflow-run-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tmpDir);
            var filePath = Path.Combine(tmpDir, "test.txt");
            File.WriteAllBytes(filePath, [0]);

            var service = new ExcelRunService();
            var request = new BridgeRequest
            {
                ProtocolVersion = ProtocolVersion.Current,
                RequestId = "req-run-open-fail",
                Command = "run",
            };
            var args = new RunCommandArguments(
                WorkbookPath: filePath,
                MacroName: "Module1.Main",
                MacroArgsJSON: "",
                Visible: false,
                DisplayAlerts: false,
                SaveWorkbook: false,
                Direct: false,
                Diagnostic: false,
                SuppressModalErrors: true,
                MsgBoxResponsesJSON: "",
                InputResponsesJSON: "",
                FileDialogResponsesJSON: "",
                DebugStreamEnabled: false,
                DebugStreamPipeName: "",
                UIStreamEnabled: false,
                UIStreamPipeName: "",
                UIStreamRedactInput: false,
                UseSession: false,
                MetadataPath: "",
                RuntimeMode: "",
                RuntimeSource: "",
                SaveAsPath: "",
                TimeoutSeconds: 0,
                ModulesDir: "",
                ClassesDir: "",
                FormsDir: "",
                WorkbookDir: "",
                CodeSource: "",
                Folders: false,
                FolderAnnotation: "update",
                DefaultComponentFolders: false);

            var response = service.Execute(request, args, CancellationToken.None);

            Assert.Equal("failed", response.Status);
            Assert.NotNull(response.Error);
            Assert.Equal("bridge_file_not_openable", response.Error.Code);
        }
        finally
        {
            if (Directory.Exists(tmpDir))
            {
                try { Directory.Delete(tmpDir, true); }
                catch (IOException) { /* best-effort */ }
            }
        }
    }

    [Fact]
    public void IsLikelyVbaCompileFailureDetectsCompileDialogHResult()
    {
        const int vbeCompileDialogHResult = unchecked((int)0x800A9C68);

        Assert.True(ExcelRunService.IsLikelyVbaCompileFailure(
            "Exception has been thrown by the target of an invocation. (0x800A9C68)",
            vbeCompileDialogHResult));
    }

    [Fact]
    public void IsLikelyVbaCompileFailureDoesNotTreatGenericRuntimeErrorAsCompile()
    {
        Assert.False(ExcelRunService.IsLikelyVbaCompileFailure("Division by zero", 11));
    }

    [Fact]
    public void IsLikelyVbaCompileFailureDoesNotTreatRuntimeSyntaxMessageAsCompile()
    {
        Assert.False(ExcelRunService.IsLikelyVbaCompileFailure("SQL syntax error near FROM clause", 5));
    }

    [Fact]
    public void WaitForPostWorkerDialogPollsCurrentDialogsWhenWatcherTaskAlreadyCompleted()
    {
        var enumerator = new SequencedWindowEnumerator(
            Array.Empty<WindowCandidate>(),
            [CompileDialogCandidate()]);
        var watcher = new DialogWatcher(enumerator, new NullUiaDialogAdapter());
        var request = new DialogWatchRequest(
            ExcelProcessId: 123,
            ExcelMainHwnd: 456,
            Kind: DialogKind.Compile,
            ActionPolicy: DialogActionPolicy.SuppressVbaError,
            Timeout: TimeSpan.FromSeconds(1),
            PollInterval: TimeSpan.FromMilliseconds(10));
        var result = new MacroRunWorkerResult(
            Completed: true,
            Ok: false,
            Value: null,
            Error: new MacroRunWorkerError("Exception has been thrown by the target of an invocation. (0x800A9C68)", "xlflow-excel-bridge", unchecked((int)0x800A9C68)));

        var dialog = ExcelRunService.WaitForPostWorkerDialog(
            watcher,
            Task.FromResult<DialogSnapshot?>(null),
            request,
            operation: "run",
            result,
            CancellationToken.None);

        Assert.NotNull(dialog);
        Assert.Equal("compile", dialog!.Kind);
        Assert.Equal("compile_close", dialog.Action);
        Assert.True(dialog.ActionSucceeded);
        Assert.Equal([11], enumerator.ClickedButtons);
    }

    [Fact]
    public void DiagnosticLocationPrefersCapturedVbeSelection()
    {
        var capture = new VbeSelectionCapture(
            new ErrorLocation(
                "high",
                "vbe.selection",
                "src/modules/Main.bas",
                "Main",
                "module",
                "Run",
                12,
                5,
                12,
                8,
                "    Debug.Print foo"),
            [new VbeSelectionCaptureAttempt("before_dialog_action", true)]);

        var location = Assert.IsType<ErrorLocation>(ExcelRunService.DiagnosticLocation(capture, new { macro = "Main.Run" }));

        Assert.Equal("src/modules/Main.bas", location.SourcePath);
        Assert.Equal("Main", location.Component);
        Assert.Equal(12, location.Line);
        Assert.Equal("vbe.selection", location.Method);
    }

    [Fact]
    public void DiagnosticLocationCapturePreservesFailedAttempts()
    {
        var capture = new VbeSelectionCapture(
            null,
            [new VbeSelectionCaptureAttempt("before_dialog_action", false, "VBE selection was not available")]);

        var metadata = ExcelRunService.DiagnosticLocationCapture(capture);
        var json = JsonSerializer.SerializeToDocument(metadata, JsonOptions.Default);

        var attempt = json.RootElement.GetProperty("attempts")[0];
        Assert.Equal("before_dialog_action", attempt.GetProperty("timing").GetString());
        Assert.False(attempt.GetProperty("success").GetBoolean());
        Assert.Contains("not available", attempt.GetProperty("error").GetString());
    }

    [Fact]
    public void VbeSelectionScorerPenalizesTemporaryHarness()
    {
        var userLocation = new ErrorLocation(
            "high",
            "vbe.selection",
            "src/modules/Main.bas",
            "Main",
            "module",
            "Run",
            6,
            3,
            6,
            8,
            "  values(1) = 2");
        var harnessLocation = userLocation with
        {
            SourcePath = "src/modules/XlflowRun_deadbeef.bas",
            Component = "XlflowRun_deadbeef",
            Procedure = "RunMacro",
        };

        Assert.True(VbeSelectionScorer.Score(userLocation) > VbeSelectionScorer.Score(harnessLocation));
        Assert.True(VbeSelectionScorer.Score(userLocation) >= VbeSelectionScorer.ReliableThreshold);
        Assert.True(VbeSelectionScorer.Score(harnessLocation) < VbeSelectionScorer.ReliableThreshold);
    }

    [Fact]
    public void VbeSelectionScorerTreatsAttributeSelectionAsUnreliable()
    {
        var location = new ErrorLocation(
            "high",
            "vbe.selection",
            "src/modules/Main.bas",
            "Main",
            "module",
            null,
            1,
            1,
            1,
            10,
            "Attribute VB_Name = \"Main\"");

        Assert.True(VbeSelectionScorer.Score(location) < VbeSelectionScorer.ReliableThreshold);
    }

    [Fact]
    public void VbeSelectionScorerTreatsOptionSelectionAsUnreliable()
    {
        var location = new ErrorLocation(
            "high",
            "vbe.selection",
            "src/modules/Main.bas",
            "Main",
            "module",
            null,
            2,
            1,
            2,
            1,
            "Option Explicit");

        Assert.True(VbeSelectionScorer.Score(location) < VbeSelectionScorer.ReliableThreshold);
    }

    [Fact]
    public void VbeSelectionScorerDetectsIncompleteCompileStatement()
    {
        Assert.True(VbeSelectionScorer.IsLikelyCompileErrorLine("  x ="));
        Assert.False(VbeSelectionScorer.IsLikelyCompileErrorLine("Attribute VB_Name = \"Main\""));
        Assert.False(VbeSelectionScorer.IsLikelyCompileErrorLine("Public Sub Run()"));
        Assert.False(VbeSelectionScorer.IsLikelyCompileErrorLine("  Dim x As String"));
    }

    [Fact]
    public void SourceLineMapperMapsVbeLineToRawSourceLine()
    {
        const string source = """
            Attribute VB_Name = "Main"
            Option Explicit

            Public Sub Run()
              Dim values(0 To 0) As Integer
              values(0) = 1
              values(1) = 2
            End Sub
            """;

        var line = SourceLineMapper.MapVbeLineToSourceLine(source, 6, "  values(1) = 2");

        Assert.Equal(7, line);
    }

    [Fact]
    public void SourceLineMapperReturnsNullWhenVbeLineCannotBeVerified()
    {
        const string source = """
            Attribute VB_Name = "Main"
            Option Explicit
            Public Sub Run()
            End Sub
            """;

        var line = SourceLineMapper.MapVbeLineToSourceLine(source, 6, "  values(1) = 2");

        Assert.Null(line);
    }

    [Fact]
    public void VbeSourcePathMapperReturnsProjectRelativePath()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-source-map-" + Guid.NewGuid().ToString("N"));
        var options = new VbeSourceMappingOptions(
            Path.Combine(root, "src", "modules"),
            Path.Combine(root, "src", "classes"),
            Path.Combine(root, "src", "forms"),
            Path.Combine(root, "src", "workbook"),
            "sidecar",
            true,
            "update",
            true);
        var path = Path.Combine(root, "src", "forms", "code", "CustomerForm.bas");

        Assert.Equal("src/forms/code/CustomerForm.bas", VbeSourcePathMapper.ToProjectRelativePath(path, options));
    }

    [Fact]
    public void VbeSourcePathMapperKeepsAbsolutePathForSiblingDirectoryOutsideProjectRoot()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-source-map-" + Guid.NewGuid().ToString("N"));
        var outsideRoot = root + "-sibling";
        var options = new VbeSourceMappingOptions(
            Path.Combine(root, "src", "modules"),
            Path.Combine(root, "src", "classes"),
            Path.Combine(root, "src", "forms"),
            Path.Combine(root, "src", "workbook"),
            "sidecar",
            true,
            "update",
            true);
        var path = Path.Combine(outsideRoot, "src", "modules", "Module1.bas");

        Assert.Equal(path, VbeSourcePathMapper.ToProjectRelativePath(path, options));
    }

    [Fact]
    public void DialogWatcherRunsPreActionCallbackBeforeClick()
    {
        var events = new List<string>();
        var enumerator = new SequencedWindowEnumerator(events, [CompileDialogCandidate()]);
        var watcher = new DialogWatcher(enumerator, new NullUiaDialogAdapter());
        var request = new DialogWatchRequest(
            ExcelProcessId: 123,
            ExcelMainHwnd: 456,
            Kind: DialogKind.Compile,
            ActionPolicy: DialogActionPolicy.SuppressVbaError,
            Timeout: TimeSpan.FromSeconds(1),
            PollInterval: TimeSpan.FromMilliseconds(10));

        _ = watcher.WaitForDialog(
            request,
            beforeAction: _ => events.Add("before"),
            afterAction: _ => events.Add("after"));

        Assert.Equal(["before", "click", "after"], events);
    }

    [Fact]
    public void DialogWatcherDoesNotRunHooksWhenActionIsNone()
    {
        var events = new List<string>();
        var enumerator = new SequencedWindowEnumerator(events, [CompileDialogCandidate()]);
        var watcher = new DialogWatcher(enumerator, new NullUiaDialogAdapter());
        var request = new DialogWatchRequest(
            ExcelProcessId: 123,
            ExcelMainHwnd: 456,
            Kind: DialogKind.Compile,
            ActionPolicy: DialogActionPolicy.ObserveOnly,
            Timeout: TimeSpan.FromSeconds(1),
            PollInterval: TimeSpan.FromMilliseconds(10));

        var dialog = watcher.WaitForDialog(
            request,
            beforeAction: _ => events.Add("before"),
            afterAction: _ => events.Add("after"));

        Assert.NotNull(dialog);
        Assert.Equal(DialogAction.None.Name, dialog!.Action);
        Assert.Empty(events);
    }

    [Fact]
    public void DialogWatcherDoesNotRunAfterHookWhenActionFails()
    {
        var events = new List<string>();
        var enumerator = new SequencedWindowEnumerator(events, false, [CompileDialogCandidate()]);
        var watcher = new DialogWatcher(enumerator, new NullUiaDialogAdapter());
        var request = new DialogWatchRequest(
            ExcelProcessId: 123,
            ExcelMainHwnd: 456,
            Kind: DialogKind.Compile,
            ActionPolicy: DialogActionPolicy.SuppressVbaError,
            Timeout: TimeSpan.FromSeconds(1),
            PollInterval: TimeSpan.FromMilliseconds(10));

        var dialog = watcher.WaitForDialog(
            request,
            beforeAction: _ => events.Add("before"),
            afterAction: _ => events.Add("after"));

        Assert.NotNull(dialog);
        Assert.False(dialog!.ActionSucceeded);
        Assert.Equal(["before", "click"], events);
    }

    [Fact]
    public void RuntimeDebugPolicyPrefersDebugButton()
    {
        var candidate = new WindowCandidate(
            Hwnd: 1,
            Pid: 123,
            ThreadId: 321,
            OwnerHwnd: 456,
            RootOwnerHwnd: 456,
            Title: "Microsoft Visual Basic",
            ClassName: "#32770",
            Visible: true,
            ProcessImage: "EXCEL.EXE",
            OwnerChain: [456],
            Text: ["Run-time error '9':", "Subscript out of range"],
            Buttons: [new WindowElement(11, "Button", "End"), new WindowElement(12, "Button", "Debug")],
            Children: []);

        var action = DialogActionSelector.Select(DialogKind.Runtime, candidate, DialogActionPolicy.SuppressVbaErrorWithRuntimeDebug);

        Assert.Equal("runtime_debug", action.Name);
        Assert.Equal(12, action.TargetHwnd);
    }

    [Fact]
    public void SaveAsExtensionIsCaseInsensitive()
    {
        ExcelRunService.AssertSaveAsExtension("C:\\work\\Book.XLSM", "C:\\work\\Copy.xlsm");
    }

    [Fact]
    public void SaveAsExtensionRejectsDifferentCaseMismatch()
    {
        Assert.Throws<InvalidOperationException>(() =>
            ExcelRunService.AssertSaveAsExtension("C:\\work\\Book.xlsm", "C:\\work\\Copy.xlsx"));
    }

    [Fact]
    public void BuildRunHarnessCodeIncludesTypedArgumentsAndErrorCapture()
    {
        var code = ExcelRunService.BuildRunHarnessCode(
            "Module1.Main",
            [
                new ExcelRunService.MacroArg { Type = "string", Value = "hello" },
                new ExcelRunService.MacroArg { Type = "int", Value = "7" },
                new ExcelRunService.MacroArg { Type = "double", Value = "3.5" },
                new ExcelRunService.MacroArg { Type = "bool", Value = "true" },
            ]);

        Assert.Contains("Application.Run targetMacro, \"hello\", CLng(7), CDbl(3.5), CBool(True)", code);
        Assert.Contains("Err.Number", code);
        Assert.Contains("Err.Description", code);
        Assert.Contains("Erl", code);
    }

    private sealed class FakeRunService(Func<BridgeRequest, RunCommandArguments, BridgeResponse> handler) : IRunService
    {
        public BridgeResponse Execute(BridgeRequest request, RunCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }

    private sealed class SequencedWindowEnumerator : IWindowEnumerator
    {
        private readonly IReadOnlyList<WindowCandidate>[] _sequences;
        private readonly List<string>? _events;
        private readonly bool _clickResult;
        private int _index;

        public SequencedWindowEnumerator(params IReadOnlyList<WindowCandidate>[] sequences)
            : this(null, true, sequences)
        {
        }

        public SequencedWindowEnumerator(List<string>? events, params IReadOnlyList<WindowCandidate>[] sequences)
            : this(events, true, sequences)
        {
        }

        public SequencedWindowEnumerator(List<string>? events, bool clickResult, params IReadOnlyList<WindowCandidate>[] sequences)
        {
            _events = events;
            _clickResult = clickResult;
            _sequences = sequences;
        }

        public List<long> ClickedButtons { get; } = [];

        public IReadOnlyList<WindowCandidate> Enumerate(int processId, int vbeThreadId)
        {
            var current = _index < _sequences.Length ? _sequences[_index] : _sequences[^1];
            if (_index < _sequences.Length - 1)
            {
                _index++;
            }
            return current;
        }

        public bool ClickButton(long hwnd)
        {
            _events?.Add("click");
            ClickedButtons.Add(hwnd);
            return _clickResult;
        }

        public bool CloseWindow(long hwnd)
        {
            return false;
        }
    }

    private sealed class NullUiaDialogAdapter : IUiaDialogAdapter
    {
        public UiaDialogDescription? Describe(long hwnd) => null;
        public bool TryInvoke(long hwnd) => false;
    }

    private static WindowCandidate CompileDialogCandidate()
    {
        return new WindowCandidate(
            Hwnd: 1,
            Pid: 123,
            ThreadId: 321,
            OwnerHwnd: 456,
            RootOwnerHwnd: 456,
            Title: "Microsoft Visual Basic for Applications",
            ClassName: "#32770",
            Visible: true,
            ProcessImage: "EXCEL.EXE",
            OwnerChain: [456],
            Text: ["コンパイル エラー:", "構文エラー"],
            Buttons: [new WindowElement(11, "Button", "OK"), new WindowElement(12, "Button", "ヘルプ")],
            Children:
            [
                new WindowElement(11, "Button", "OK"),
                new WindowElement(12, "Button", "ヘルプ"),
                new WindowElement(13, "Static", "コンパイル エラー:"),
                new WindowElement(14, "Static", "構文エラー"),
            ]);
    }
}
