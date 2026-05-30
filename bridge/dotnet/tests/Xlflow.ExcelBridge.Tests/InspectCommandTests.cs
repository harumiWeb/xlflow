using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class InspectCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeInspectService((request, args) =>
        {
            Assert.Equal("inspect", request.Command);
            Assert.Equal("range", args.Target);
            Assert.Equal("Sheet1", args.Sheet);
            Assert.Equal("A1:B2", args.Address);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);
            Assert.True(args.UseSession);
            Assert.True(args.IncludeStyle);
            Assert.Equal(25, args.MaxRows);
            Assert.Equal(8, args.MaxCols);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?>
                {
                    ["kind"] = "live_session",
                    ["path"] = args.WorkbookPath,
                },
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = true,
                    ["workbook_path"] = args.WorkbookPath,
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = args.WorkbookPath,
                    ["session"] = true,
                },
                ["inspect"] = new Dictionary<string, object?>
                {
                    ["target"] = args.Target,
                    ["source"] = "excel_com",
                    ["range"] = new Dictionary<string, object?>
                    {
                        ["sheet"] = args.Sheet,
                        ["range"] = "$A$1:$B$2",
                        ["returned_range"] = "$A$1:$B$2",
                        ["row_count"] = 2,
                        ["column_count"] = 2,
                        ["values"] = new object?[][]
                        {
                            ["a", "b"],
                            ["c", "d"],
                        },
                        ["style_included"] = true,
                        ["merged_ranges"] = Array.Empty<string>(),
                    },
                },
            });
        });
        var command = new InspectCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-inspect",
            Command = "inspect",
            Payload = JsonDocument.Parse("""
                {
                  "Target": "range",
                  "Sheet": "Sheet1",
                  "Address": "A1:B2",
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json",
                  "UseSession": "true",
                  "IncludeStyle": "true",
                  "MaxRows": "25",
                  "MaxCols": "8"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("range", json.RootElement.GetProperty("inspect").GetProperty("target").GetString());
        Assert.Equal("excel_com", json.RootElement.GetProperty("inspect").GetProperty("source").GetString());
        Assert.Equal("$A$1:$B$2", json.RootElement.GetProperty("inspect").GetProperty("range").GetProperty("returned_range").GetString());
        Assert.True(json.RootElement.GetProperty("inspect").GetProperty("range").GetProperty("style_included").GetBoolean());
    }

    [Fact]
    public void HandleRejectsMissingTarget()
    {
        var command = new InspectCommand(new FakeInspectService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-inspect-missing-target",
            Command = "inspect",
            Payload = JsonDocument.Parse("""{"WorkbookPath":"C:\\work\\book.xlsm"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("inspect_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeInspectService(Func<BridgeRequest, InspectCommandArguments, BridgeResponse> handler) : IInspectService
    {
        public BridgeResponse Execute(BridgeRequest request, InspectCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }
}
