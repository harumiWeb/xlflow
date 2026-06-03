using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

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
            ],
            traceEnabled: false,
            traceFile: "");

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
}
