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
            new DoctorCommand(),
            new InspectCommand(),
            new MacrosCommand(),
            new ProcessCommand(),
            new PullCommand(),
            new PushCommand(),
            new RunCommand(),
        });
    }

    public static CommandRegistry Create(
        Func<ExcelDiagnosticsResult>? probeExcel,
        IInspectService? inspectService = null,
        IMacrosService? macrosService = null,
        IProcessService? processService = null,
        IPullService? pullService = null,
        IPushService? pushService = null,
        IRunService? runService = null)
    {
        return new CommandRegistry(new ICommandHandler[]
        {
            new DoctorCommand(probeExcel),
            new InspectCommand(inspectService),
            new MacrosCommand(macrosService),
            new ProcessCommand(processService),
            new PullCommand(pullService),
            new PushCommand(pushService),
            new RunCommand(runService),
        });
    }

    public ICommandHandler? Resolve(string command)
    {
        return _handlers.TryGetValue(command, out var handler) ? handler : null;
    }
}
