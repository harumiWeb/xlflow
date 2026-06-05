using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class FormWriteCommand : ICommandHandler
{
    private readonly IFormWriteService _service;

    public FormWriteCommand(IFormWriteService? service = null)
    {
        _service = service ?? new ExcelFormWriteService();
    }

    public string CommandName => "form-write";

    public bool Supports(BridgeRequest request)
    {
        return string.Equals(request.Command, CommandName, StringComparison.OrdinalIgnoreCase);
    }

    public BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken)
    {
        var args = new FormWriteCommandArguments(
            Action: BridgePayload.GetString(request.Payload, "Action") ?? "",
            WorkbookPath: BridgePayload.GetString(request.Payload, "WorkbookPath") ?? "",
            SpecPath: BridgePayload.GetString(request.Payload, "SpecPath") ?? "",
            FormsDir: BridgePayload.GetString(request.Payload, "FormsDir") ?? "",
            CodeSource: BridgePayload.GetString(request.Payload, "CodeSource") ?? "frm",
            Folders: BridgePayload.GetBool(request.Payload, "Folders"),
            FolderAnnotation: BridgePayload.GetString(request.Payload, "FolderAnnotation") ?? "update",
            DefaultComponentFolders: BridgePayload.GetBool(request.Payload, "DefaultComponentFolders"),
            SpecJson64: BridgePayload.GetString(request.Payload, "SpecJson64") ?? "",
            Overwrite: BridgePayload.GetBool(request.Payload, "Overwrite"),
            NoSave: BridgePayload.GetBool(request.Payload, "NoSave"),
            Visible: BridgePayload.GetBool(request.Payload, "Visible"),
            UseSession: BridgePayload.GetBool(request.Payload, "UseSession"),
            MetadataPath: BridgePayload.GetString(request.Payload, "MetadataPath") ?? "");

        var action = args.Action.Trim().ToLowerInvariant();
        var validationCode = action == "apply" ? "form_apply_args_invalid" : "form_build_args_invalid";

        if (action is not ("build" or "apply"))
        {
            return Failed(request, validationCode, $"unsupported form action: {args.Action}");
        }
        if (string.IsNullOrWhiteSpace(args.WorkbookPath))
        {
            return Failed(request, validationCode, "WorkbookPath is required.");
        }
        if (string.IsNullOrWhiteSpace(args.SpecJson64))
        {
            return Failed(request, validationCode, "SpecJson64 is required.");
        }
        if (string.IsNullOrWhiteSpace(args.FormsDir))
        {
            return Failed(request, validationCode, "FormsDir is required.");
        }
        if (args.NoSave && !args.UseSession)
        {
            return Failed(request, validationCode, "--NoSave requires --UseSession.");
        }
        if (action == "build" && args.Overwrite && args.NoSave)
        {
            return Failed(request, validationCode, "--overwrite cannot be combined with --NoSave because Excel requires an intermediate save before recreating the UserForm.");
        }

        return _service.Execute(request, args with { Action = action }, cancellationToken);
    }

    private static BridgeResponse Failed(BridgeRequest request, string code, string message)
    {
        return BridgeResponse.Failed(request, new BridgeError(
            Code: code,
            Message: message,
            Phase: "validate_args",
            Source: "xlflow"));
    }
}
