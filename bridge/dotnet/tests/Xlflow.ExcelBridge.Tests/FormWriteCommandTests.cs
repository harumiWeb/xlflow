using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class FormWriteCommandTests
{
    [Theory]
    [InlineData("prepare", "PreparedWorkbookPath")]
    [InlineData("import", "CompiledWorkbookPath")]
    public void HandleRequiresPrivateWorkbookPathForSelectedPhase(string phase, string pathField)
    {
        var service = new FakeFormWriteService((request, _) => BridgeResponse.Ok(request));
        var command = new FormWriteCommand(service);
        var request = BuildRequest($$"""
            {
              "Action": "build",
              "WorkbookPath": "C:\\work\\book.xlsm",
              "SpecJson64": "eyJmb3JtIjp7Im5hbWUiOiJTYW1wbGUifX0=",
              "Phase": "{{phase}}"
            }
            """);

        var response = command.Handle(request, CancellationToken.None);

        Assert.Equal("failed", response.Status);
        Assert.Equal("form_build_args_invalid", response.Error!.Code);
        Assert.Contains(pathField, response.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HandlePassesPrepareAndImportArgumentsWithoutRequiringFormSpecMirrors()
    {
        var calls = new List<FormWriteCommandArguments>();
        var service = new FakeFormWriteService((request, args) =>
        {
            calls.Add(args);
            return BridgeResponse.Ok(request);
        });
        var command = new FormWriteCommand(service);
        foreach (var phase in new[] { "prepare", "import" })
        {
            var phasePath = phase == "prepare" ? "PreparedWorkbookPath" : "CompiledWorkbookPath";
            var request = BuildRequest($$"""
                {
                  "Action": "apply",
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "SpecPath": "src\\forms\\Sample.yaml",
                  "SpecJson64": "eyJmb3JtIjp7Im5hbWUiOiJTYW1wbGUifX0=",
                  "Phase": "{{phase}}",
                  "{{phasePath}}": "C:\\temp\\stage.xlsm",
                  "NoSave": "true",
                  "UseSession": "true"
                }
                """);

            Assert.Equal("ok", command.Handle(request, CancellationToken.None).Status);
        }

        Assert.Collection(calls,
            prepare =>
            {
                Assert.Equal("prepare", prepare.Phase);
                Assert.Equal(@"C:\temp\stage.xlsm", prepare.PreparedWorkbookPath);
                Assert.True(prepare.NoSave);
            },
            import =>
            {
                Assert.Equal("import", import.Phase);
                Assert.Equal(@"C:\temp\stage.xlsm", import.CompiledWorkbookPath);
                Assert.True(import.NoSave);
            });
    }

    [Fact]
    public void HandleRejectsNoSaveWithoutExplicitSession()
    {
        var command = new FormWriteCommand(new FakeFormWriteService((request, _) => BridgeResponse.Ok(request)));
        var request = BuildRequest("""
            {
              "Action": "build",
              "WorkbookPath": "C:\\work\\book.xlsm",
              "SpecJson64": "eyJmb3JtIjp7Im5hbWUiOiJTYW1wbGUifX0=",
              "Phase": "prepare",
              "PreparedWorkbookPath": "C:\\temp\\stage.xlsm",
              "NoSave": "true",
              "UseSession": "false"
            }
            """);

        var response = command.Handle(request, CancellationToken.None);

        Assert.Equal("failed", response.Status);
        Assert.Equal("form_build_args_invalid", response.Error!.Code);
    }

    [Fact]
    public void ReplaceFormWithRollbackRestoresExportedExistingFormWhenImportFails()
    {
        var calls = new List<string>();

        var error = Assert.Throws<InvalidOperationException>(() => ExcelFormWriteService.ReplaceFormWithRollback(
            hadExistingForm: true,
            removeExisting: () => calls.Add("remove-existing"),
            importCompiled: () => throw new InvalidOperationException("compiled import failed"),
            removePartialImport: () => calls.Add("remove-partial"),
            restoreExisting: () => calls.Add("restore-backup")));

        Assert.Equal("compiled import failed", error.Message);
        Assert.Equal(new[] { "remove-existing", "remove-partial", "restore-backup" }, calls);
    }

    [Fact]
    public void ReplaceFormWithRollbackRemovesPartialNewFormWhenImportFails()
    {
        var calls = new List<string>();

        _ = Assert.Throws<InvalidOperationException>(() => ExcelFormWriteService.ReplaceFormWithRollback(
            hadExistingForm: false,
            removeExisting: () => calls.Add("unexpected-remove-existing"),
            importCompiled: () => throw new InvalidOperationException("compiled import failed"),
            removePartialImport: () => calls.Add("remove-partial"),
            restoreExisting: () => calls.Add("unexpected-restore")));

        Assert.Equal(new[] { "remove-partial" }, calls);
    }

    [Fact]
    public void ReplaceFormWithRollbackReportsRollbackFailureAlongsideImportFailure()
    {
        var error = Assert.Throws<InvalidOperationException>(() => ExcelFormWriteService.ReplaceFormWithRollback(
            hadExistingForm: true,
            removeExisting: () => { },
            importCompiled: () => throw new InvalidOperationException("compiled import failed"),
            removePartialImport: () => throw new InvalidOperationException("remove failed"),
            restoreExisting: () => { }));

        Assert.Contains("rollback failed", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.IsType<AggregateException>(error.InnerException);
    }

    private static BridgeRequest BuildRequest(string payload)
    {
        return new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-write",
            Command = "form-write",
            Payload = JsonDocument.Parse(payload).RootElement.Clone(),
        };
    }

    private sealed class FakeFormWriteService(Func<BridgeRequest, FormWriteCommandArguments, BridgeResponse> handler) : IFormWriteService
    {
        public BridgeResponse Execute(BridgeRequest request, FormWriteCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
