using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;

namespace Xlflow.ExcelBridge;

public static class BridgeHost
{
    [SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "The bridge host must translate unexpected failures into structured JSON responses on stdout.")]
    public static int Run(string[] args, TextReader stdin, TextWriter stdout, TextWriter stderr)
    {
        try
        {
            var registry = CommandRegistry.CreateDefault();

            if (args.Contains("--version-json", StringComparer.OrdinalIgnoreCase))
            {
                StdinStdoutTransport.Write(stdout, BridgeInfo.Current());
                return 0;
            }

            if (args.Contains("--capabilities-json", StringComparer.OrdinalIgnoreCase))
            {
                StdinStdoutTransport.Write(stdout, new BridgeCapabilities(registry.CommandNames));
                return 0;
            }

            var request = StdinStdoutTransport.Read<BridgeRequest>(stdin);
            if (request is null)
            {
                StdinStdoutTransport.Write(stdout, BridgeResponse.Failed(
                    BridgeRequest.Unknown,
                    BridgeError.Create("BRIDGE_REQUEST_EMPTY", "No bridge request JSON was provided.", "transport.read")));
                return 2;
            }

            var handler = registry.Resolve(request.Command);
            if (handler is null || !handler.Supports(request))
            {
                StdinStdoutTransport.Write(stdout, BridgeResponse.Failed(
                    request,
                    BridgeError.Create("BRIDGE_COMMAND_UNSUPPORTED", $"Command '{request.Command}' is not supported by the .NET bridge.", "bridge.capability")));
                return 3;
            }

            var response = handler.Handle(request, CancellationToken.None);
            StdinStdoutTransport.Write(stdout, response);
            return response.Status == BridgeStatus.Ok ? 0 : 1;
        }
        catch (JsonException ex)
        {
            StdinStdoutTransport.Write(stdout, BridgeResponse.Failed(
                BridgeRequest.Unknown,
                BridgeError.Create("BRIDGE_REQUEST_INVALID_JSON", ex.Message, "transport.read")));
            return 2;
        }
        catch (Exception ex)
        {
            stderr.WriteLine(ex);
            StdinStdoutTransport.Write(stdout, BridgeResponse.Failed(
                BridgeRequest.Unknown,
                BridgeError.Create("BRIDGE_FATAL", ex.Message, "bridge.host")));
            return 1;
        }
    }
}
