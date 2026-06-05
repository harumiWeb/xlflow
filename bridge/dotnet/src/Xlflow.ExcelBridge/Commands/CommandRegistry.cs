using Xlflow.ExcelBridge.Diagnostics;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Commands;

public sealed class CommandRegistry
{
    private readonly Dictionary<string, ICommandHandler> _handlers;

    private CommandRegistry(IEnumerable<ICommandHandler> handlers)
    {
        _handlers = handlers.ToDictionary(
            handler => handler.CommandName,
            handler => handler,
            StringComparer.OrdinalIgnoreCase);
    }

    public IReadOnlyList<string> CommandNames => _handlers.Keys.Order(StringComparer.OrdinalIgnoreCase).ToArray();

    public static CommandRegistry CreateDefault()
    {
        return new CommandRegistry(new ICommandHandler[]
        {
            new AttachCommand(),
            new DoctorCommand(),
            new EditCommand(),
            new ExportImageCommand(),
            new FormExportImageCommand(),
            new FormWriteCommand(),
            new InspectCommand(),
            new InspectFormCommand(),
            new ListCommand(),
            new MacrosCommand(),
            new NewCommand(),
            new ProcessCommand(),
            new PullCommand(),
            new PushCommand(),
            new RunCommand(),
            new RunnerCommand(),
            new SessionCommand(),
            new TestCommand(),
            new TraceCommand(),
            new UICommand(),
        });
    }

    public static CommandRegistry Create(
        Func<ExcelDiagnosticsResult>? probeExcel,
        IAttachService? attachService = null,
        IEditService? editService = null,
        IExportImageService? exportImageService = null,
        IFormExportImageService? formExportImageService = null,
        IInspectService? inspectService = null,
        IFormWriteService? formWriteService = null,
        IInspectFormService? inspectFormService = null,
        IListService? listService = null,
        IMacrosService? macrosService = null,
        INewService? newService = null,
        IProcessService? processService = null,
        IPullService? pullService = null,
        IPushService? pushService = null,
        IRunService? runService = null,
        IRunnerService? runnerService = null,
        ISessionService? sessionService = null,
        ITestService? testService = null,
        ITraceService? traceService = null,
        IUIService? uiService = null)
    {
        return new CommandRegistry(new ICommandHandler[]
        {
            new AttachCommand(attachService),
            new DoctorCommand(probeExcel),
            new EditCommand(editService),
            new ExportImageCommand(exportImageService),
            new FormExportImageCommand(formExportImageService),
            new FormWriteCommand(formWriteService),
            new InspectCommand(inspectService),
            new InspectFormCommand(inspectFormService),
            new ListCommand(listService),
            new MacrosCommand(macrosService),
            new NewCommand(newService),
            new ProcessCommand(processService),
            new PullCommand(pullService),
            new PushCommand(pushService),
            new RunCommand(runService),
            new RunnerCommand(runnerService),
            new SessionCommand(sessionService),
            new TestCommand(testService),
            new TraceCommand(traceService),
            new UICommand(uiService),
        });
    }

    public ICommandHandler? Resolve(string command)
    {
        return _handlers.TryGetValue(command, out var handler) ? handler : null;
    }
}
