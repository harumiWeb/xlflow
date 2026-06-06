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
                  "DebugStreamPipeName": "\\\\.\\pipe\\xlflow-debug-test"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.True(serviceCalled);
        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
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
                TimeoutSeconds: 0);

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
            [],
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

    private sealed class SequencedWindowEnumerator(params IReadOnlyList<WindowCandidate>[] sequences) : IWindowEnumerator
    {
        private int _index;

        public List<long> ClickedButtons { get; } = [];

        public IReadOnlyList<WindowCandidate> Enumerate(int processId, int vbeThreadId)
        {
            var current = _index < sequences.Length ? sequences[_index] : sequences[^1];
            if (_index < sequences.Length - 1)
            {
                _index++;
            }
            return current;
        }

        public bool ClickButton(long hwnd)
        {
            ClickedButtons.Add(hwnd);
            return true;
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
