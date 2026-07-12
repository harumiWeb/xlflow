using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class InspectFormCommandTests
{
    [Theory]
    [InlineData("Forms.Label.1", "__ComObject", "Label")]
    [InlineData("Forms.TextBox.1", "__ComObject", "TextBox")]
    [InlineData("", "__ComObject", "__ComObject")]
    [InlineData(null, "Frame", "Frame")]
    public void ResolveDesignerControlType_PrefersProgIdSegment(string? progId, string fallbackTypeName, string expected)
    {
        var actual = ExcelFormInspectionService.ResolveDesignerControlType(progId, fallbackTypeName);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("MSForms.TextBox", "TextBox")]
    [InlineData("CommandButtonClass", "CommandButton")]
    [InlineData("IMdcCombo", "ComboBox")]
    [InlineData("IMdcText", "TextBox")]
    [InlineData("ILabelControl", "Label")]
    [InlineData("IOptionFrame", "Frame")]
    [InlineData("ICommandButton", "CommandButton")]
    [InlineData("_Label", "Label")]
    [InlineData("_IMdcText", "TextBox")]
    [InlineData("_IMdcTextClass", "TextBox")]
    [InlineData("__ComObject", null)]
    [InlineData("", null)]
    public void NormalizeComTypeName_CleansDesignerComNames(string? value, string? expected)
    {
        var actual = ExcelFormInspectionService.NormalizeComTypeName(value);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void GetChildControls_TreatsControlsPropertyMissingAsEmpty()
    {
        var controls = ExcelFormInspectionService.GetChildControls(new DesignerControlWithoutChildren(), "Parent");

        Assert.Empty(controls);
    }

    [Fact]
    public void GetChildControls_PropagatesEnumerationFailures()
    {
        Assert.ThrowsAny<Exception>(() => ExcelFormInspectionService.GetChildControls(new DesignerControlWithBrokenChildren(), "Parent"));
    }

    [Fact]
    public void TryGetDesignerFormDimension_PrefersComponentProperty()
    {
        var actual = ExcelFormInspectionService.TryGetDesignerFormDimension(
            new DesignerComponentWithProperties(),
            new DesignerFormWithDimensions(width: 0, height: 0),
            "Width");

        Assert.Equal(300.0, actual);
    }

    [Fact]
    public void TryGetDesignerFormDimension_FallsBackToDesignerMember()
    {
        var actual = ExcelFormInspectionService.TryGetDesignerFormDimension(
            new DesignerComponentWithoutProperties(),
            new DesignerFormWithDimensions(width: 280.5, height: 240.25),
            "Height");

        Assert.Equal(240.25, actual);
    }

    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeInspectFormService((request, args) =>
        {
            Assert.Equal("inspect-form", request.Command);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal("UserForm1", args.FormName);
            Assert.Equal("both", args.Basis);
            Assert.Equal("InitializeRuntime", args.Initializer);
            Assert.True(args.StrictDesigner);
            Assert.False(args.Visible);
            Assert.True(args.UseSession);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?>
                {
                    ["kind"] = "live_session",
                    ["path"] = args.WorkbookPath,
                    ["note"] = "Runtime inspection used a temporary workbook copy.",
                },
                ["session"] = new Dictionary<string, object?>
                {
                    ["active"] = true,
                    ["workbook_path"] = args.WorkbookPath,
                    ["dirty"] = false,
                    ["save_required"] = false,
                    ["mode"] = "explicit",
                },
                ["workbook"] = new Dictionary<string, object?>
                {
                    ["path"] = args.WorkbookPath,
                    ["session"] = true,
                    ["session_mode"] = "explicit",
                },
                ["forms"] = new Dictionary<string, object?>
                {
                    ["runtime"] = new Dictionary<string, object?>
                    {
                        ["name"] = args.FormName,
                        ["basis"] = "runtime",
                    },
                    ["designer"] = new Dictionary<string, object?>
                    {
                        ["name"] = args.FormName,
                        ["basis"] = "designer",
                    },
                },
                ["warnings"] = new List<Dictionary<string, string>>
                {
                    new() { ["code"] = "runtime_form_temp_copy", ["message"] = "Runtime inspection executed against a temporary workbook copy so the source workbook and live session are not mutated." },
                },
            });
        });
        var command = new InspectFormCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-inspect-form",
            Command = "inspect-form",
            Payload = JsonDocument.Parse("""
                {
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "FormName": "UserForm1",
                  "Basis": "both",
                  "Initializer": "InitializeRuntime",
                  "StrictDesigner": "true",
                  "Visible": "false",
                  "UseSession": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("live_session", json.RootElement.GetProperty("target").GetProperty("kind").GetString());
        Assert.Equal("runtime", json.RootElement.GetProperty("forms").GetProperty("runtime").GetProperty("basis").GetString());
        Assert.Equal("designer", json.RootElement.GetProperty("forms").GetProperty("designer").GetProperty("basis").GetString());
        Assert.Equal("runtime_form_temp_copy", json.RootElement.GetProperty("warnings")[0].GetProperty("code").GetString());
    }

    [Fact]
    public void HandleRejectsMissingWorkbookPath()
    {
        var command = new InspectFormCommand(new FakeInspectFormService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-inspect-form-missing-workbook",
            Command = "inspect-form",
            Payload = JsonDocument.Parse("""{"FormName":"UserForm1"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("inspect_form_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void HandleRejectsMissingFormName()
    {
        var command = new InspectFormCommand(new FakeInspectFormService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-inspect-form-missing-form",
            Command = "inspect-form",
            Payload = JsonDocument.Parse("""{"WorkbookPath":"C:\\work\\book.xlsm"}""").RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("inspect_form_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    private sealed class FakeInspectFormService(Func<BridgeRequest, InspectFormCommandArguments, BridgeResponse> handler) : IInspectFormService
    {
        public BridgeResponse Execute(BridgeRequest request, InspectFormCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }

    private sealed class DesignerControlWithoutChildren
    {
    }

    private sealed class DesignerControlWithBrokenChildren
    {
        public BrokenControls Controls => new();
    }

    private sealed class DesignerFormWithDimensions(double width, double height)
    {
        public double Width { get; } = width;

        public double Height { get; } = height;
    }

    private sealed class DesignerComponentWithProperties
    {
        public DesignerProperties Properties { get; } = new();
    }

    private sealed class DesignerComponentWithoutProperties
    {
    }

    private sealed class DesignerProperties
    {
        public DesignerProperty Item(string name)
        {
            return name switch
            {
                "Width" => new DesignerProperty(300.0),
                "Height" => new DesignerProperty(262.2),
                _ => throw new ArgumentException("unknown designer property", nameof(name)),
            };
        }
    }

    private sealed class DesignerProperty(double value)
    {
        public double Value { get; } = value;
    }

    private sealed class BrokenControls
    {
        public int Count => throw new InvalidOperationException("broken child enumeration");
    }
}
