using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class EditCommandTests
{
    [Fact]
    public void HandleParsesCellEditArguments()
    {
        var service = new FakeEditService((request, args) =>
        {
            Assert.Equal("edit", request.Command);
            Assert.Equal("cell", args.Action);
            Assert.Equal(@"C:\work\Book.xlsm", args.WorkbookPath);
            Assert.Equal("Sheet1", args.Sheet);
            Assert.Equal("C3", args.Cell);
            Assert.Equal("42", args.Value);
            Assert.True(args.ValueSpecified);
            Assert.False(args.FormulaSpecified);
            Assert.False(args.FormulaR1C1Specified);
            Assert.Equal("off", args.Events);
            Assert.True(args.UseSession);
            Assert.Equal(@".xlflow\session.json", args.MetadataPath);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>());
        });

        var command = new EditCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-edit",
            Command = "edit",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath":"C:\\work\\Book.xlsm",
                  "Action":"cell",
                  "Sheet":"Sheet1",
                  "Cell":"C3",
                  "Value":"42",
                  "Events":"off",
                  "UseSession":"true",
                  "MetadataPath":".xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new EditCommand(new FakeEditService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-edit-bad",
            Command = "edit",
            Payload = JsonDocument.Parse("""{"Action":"cell"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);
        Assert.Equal("edit_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandlePreservesExplicitEmptyValue()
    {
        var service = new FakeEditService((request, args) =>
        {
            Assert.Equal(string.Empty, args.Value);
            Assert.True(args.ValueSpecified);
            Assert.False(args.FormulaSpecified);
            Assert.False(args.FormulaR1C1Specified);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>());
        });

        var command = new EditCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-edit-empty",
            Command = "edit",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath":"C:\\work\\Book.xlsm",
                  "Action":"cell",
                  "Sheet":"Sheet1",
                  "Cell":"C3",
                  "Value":"",
                  "UseSession":"true",
                  "MetadataPath":".xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandleParsesFormulaEditArguments()
    {
        var service = new FakeEditService((request, args) =>
        {
            Assert.Equal("formula", args.Action);
            Assert.Equal("Invoice", args.Sheet);
            Assert.Equal("D2:D1001", args.RangeAddress);
            Assert.Equal("=RC[-2]*RC[-1]", args.FormulaR1C1);
            Assert.True(args.FormulaR1C1Specified);
            Assert.False(args.FormulaSpecified);
            Assert.Equal("off", args.Events);
            Assert.True(args.Calculate);
            Assert.True(args.UseSession);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>());
        });

        var command = new EditCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-edit-formula",
            Command = "edit",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath":"C:\\work\\Book.xlsm",
                  "Action":"formula",
                  "Sheet":"Invoice",
                  "RangeAddress":"D2:D1001",
                  "FormulaR1C1":"=RC[-2]*RC[-1]",
                  "Events":"off",
                  "Calculate":"true",
                  "UseSession":"true",
                  "MetadataPath":".xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void HandlePreservesExplicitEmptyFormula()
    {
        var service = new FakeEditService((request, args) =>
        {
            Assert.Equal("formula", args.Action);
            Assert.Equal(string.Empty, args.Formula);
            Assert.True(args.FormulaSpecified);
            Assert.False(args.FormulaR1C1Specified);
            return BridgeResponse.Ok(request, new Dictionary<string, object?>());
        });

        var command = new EditCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-edit-empty-formula",
            Command = "edit",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath":"C:\\work\\Book.xlsm",
                  "Action":"formula",
                  "Sheet":"Invoice",
                  "RangeAddress":"D2",
                  "Formula":"",
                  "UseSession":"true",
                  "MetadataPath":".xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        Assert.Equal("ok", JsonSerializer.SerializeToDocument(response, JsonOptions.Default).RootElement.GetProperty("status").GetString());
    }

    private sealed class FakeEditService(Func<BridgeRequest, EditCommandArguments, BridgeResponse> handler) : IEditService
    {
        public BridgeResponse Execute(BridgeRequest request, EditCommandArguments args, CancellationToken cancellationToken) => handler(request, args);
    }
}
