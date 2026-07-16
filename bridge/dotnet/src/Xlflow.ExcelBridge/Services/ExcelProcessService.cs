using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Services;

[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "The .NET bridge only runs on Windows with Excel COM automation.")]
[SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Process cleanup/list normalize OS and COM failures into structured bridge responses.")]
[SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance", Justification = "The current signatures favor testability and concise cleanup logic.")]
public sealed class ExcelProcessService : IProcessService
{
    public BridgeResponse Execute(BridgeRequest request, ProcessCommandArguments args, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            return args.Action switch
            {
                "list" => List(request, args),
                "cleanup" => Cleanup(request, args),
                _ => BridgeResponse.Failed(request, new BridgeError("process_args_invalid", "Action must be list or cleanup", "process", "xlflow")),
            };
        }
        catch (Exception ex)
        {
            return BridgeResponse.Failed(request, new BridgeError(
                Code: "process_enumeration_failed",
                Message: ex.Message,
                Phase: "process",
                Source: "xlflow-excel-bridge"));
        }
    }

    private static BridgeResponse List(BridgeRequest request, ProcessCommandArguments args)
    {
        var processes = ExcelBridgeSupport.GetExcelProcesses(args.SkipWorkbookProbePids)
            .OrderBy(item => item.ProcessId)
            .Select(item => BuildProcessListItem(item, args.SkipWorkbookProbePids))
            .ToArray();

        var logs = processes.Length == 0
            ? new[] { "0 Excel processes found" }
            : new[] { $"found {processes.Length} Excel process(es)" };

        return new BridgeResponse
        {
            RequestId = request.RequestId,
            Command = request.Command,
            Logs = logs,
            Extensions = new Dictionary<string, object?>
            {
                ["process"] = processes,
            },
        };
    }

    internal static Dictionary<string, object?> BuildProcessListItem(
        ExcelProcessInfo process,
        IReadOnlySet<int> skipWorkbookProbePids)
    {
        var recoveryRequired = skipWorkbookProbePids.Contains(process.ProcessId);
        return new Dictionary<string, object?>
        {
            ["pid"] = process.ProcessId,
            ["has_workbook"] = process.HasWorkbook,
            ["workbook_probe_skipped"] = recoveryRequired,
            ["recovery_required"] = recoveryRequired,
        };
    }

    private static BridgeResponse Cleanup(BridgeRequest request, ProcessCommandArguments args)
    {
        var mode = "";
        var targetPids = new List<int>();
        var targetProcesses = new Dictionary<int, Process>();

        try
        {
            if (args.TargetPid is > 0)
            {
                mode = "pid";
                Process? process = null;
                try
                {
                    process = Process.GetProcessById(args.TargetPid.Value);
                    if (!string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                    {
                        return ProcessNotFound(request, args.TargetPid.Value);
                    }

                    targetPids.Add(args.TargetPid.Value);
                }
                catch (ArgumentException)
                {
                    process?.Dispose();
                    return ProcessNotFound(request, args.TargetPid.Value);
                }
                finally
                {
                    process?.Dispose();
                }
            }
            else if (args.Auto)
            {
                mode = "auto";
                foreach (var process in ExcelBridgeSupport.GetExcelProcesses())
                {
                    if (process.HasWorkbook == false)
                    {
                        targetPids.Add(process.ProcessId);
                    }
                }
            }
            else if (args.All)
            {
                mode = "all";
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    targetPids.Add(process.Id);
                    targetProcesses[process.Id] = process;
                }
            }
            else
            {
                return BridgeResponse.Failed(request, new BridgeError(
                    Code: "process_args_invalid",
                    Message: "process cleanup requires a PID, --auto, or --all",
                    Phase: "process",
                    Source: "xlflow"));
            }

            if (targetPids.Count == 0)
            {
                return new BridgeResponse
                {
                    RequestId = request.RequestId,
                    Command = request.Command,
                    Logs = ["0 Excel processes to clean up"],
                    Extensions = new Dictionary<string, object?>
                    {
                        ["process"] = new Dictionary<string, object?>
                        {
                            ["action"] = "cleanup",
                            ["mode"] = mode,
                            ["total"] = 0,
                            ["results"] = Array.Empty<object>(),
                        },
                    },
                };
            }

            var cleanupResults = new List<Dictionary<string, object?>>();
            foreach (var targetPid in targetPids)
            {
                cleanupResults.Add(args.All
                    ? StopAllTarget(targetProcesses, targetPid)
                    : StopProcessTarget(targetPid));
            }

            var terminatedCount = cleanupResults.Count(item => item.TryGetValue("terminated", out var value) && value is true);
            var failedCount = cleanupResults.Count - terminatedCount;
            var processPayload = new Dictionary<string, object?>
            {
                ["action"] = "cleanup",
                ["mode"] = mode,
                ["total"] = cleanupResults.Count,
                ["results"] = cleanupResults,
            };

            if (failedCount > 0)
            {
                return new BridgeResponse
                {
                    RequestId = request.RequestId,
                    Command = request.Command,
                    Status = BridgeStatus.Failed,
                    Error = new BridgeError(
                        Code: "process_termination_failed",
                        Message: $"{failedCount} of {cleanupResults.Count} Excel process(es) failed to terminate",
                        Phase: "process",
                        Source: "xlflow"),
                    Extensions = new Dictionary<string, object?>
                    {
                        ["process"] = processPayload,
                    },
                };
            }

            return new BridgeResponse
            {
                RequestId = request.RequestId,
                Command = request.Command,
                Logs = [$"terminated {terminatedCount} Excel process(es)"],
                Extensions = new Dictionary<string, object?>
                {
                    ["process"] = processPayload,
                },
            };
        }
        finally
        {
            foreach (var process in targetProcesses.Values)
            {
                process.Dispose();
            }
        }
    }

    private static BridgeResponse ProcessNotFound(BridgeRequest request, int pid)
    {
        return BridgeResponse.Failed(request, new BridgeError(
            Code: "process_not_found",
            Message: $"no Excel process found with PID {pid}",
            Phase: "process",
            Source: "xlflow"));
    }

    private static Dictionary<string, object?> StopAllTarget(Dictionary<int, Process> targetProcesses, int targetPid)
    {
        if (!targetProcesses.TryGetValue(targetPid, out var process))
        {
            return NewCleanupResult(targetPid, true, "unknown");
        }

        try
        {
            process.Refresh();
            if (process.HasExited)
            {
                return NewCleanupResult(targetPid, true, "unknown");
            }

            process.Kill(true);
            var exited = process.WaitForExit(3000);
            process.Refresh();
            if (!exited && !process.HasExited)
            {
                return NewCleanupResult(targetPid, false, "force");
            }
            return NewCleanupResult(targetPid, true, "force");
        }
        catch
        {
            return NewCleanupResult(targetPid, false, "none");
        }
    }

    private static Dictionary<string, object?> StopProcessTarget(int targetPid)
    {
        try
        {
            using var process = Process.GetProcessById(targetPid);
            return StopProcess(process);
        }
        catch (ArgumentException)
        {
            return NewCleanupResult(targetPid, true, "unknown");
        }
        catch
        {
            return NewCleanupResult(targetPid, false, "none");
        }
    }

    private static Dictionary<string, object?> StopProcess(Process process)
    {
        try
        {
            try
            {
                if (process.CloseMainWindow())
                {
                    if (process.WaitForExit(3000))
                    {
                        return NewCleanupResult(process.Id, true, "graceful");
                    }
                }
            }
            catch
            {
                // fall through to force termination
            }

            process.Refresh();
            if (process.HasExited)
            {
                return NewCleanupResult(process.Id, true, "unknown");
            }

            process.Kill(true);
            var exited = process.WaitForExit(3000);
            process.Refresh();
            if (!exited && !process.HasExited)
            {
                return NewCleanupResult(process.Id, false, "force");
            }
            return NewCleanupResult(process.Id, true, "force");
        }
        catch
        {
            return NewCleanupResult(process.Id, false, "none");
        }
    }

    private static Dictionary<string, object?> NewCleanupResult(int pid, bool terminated, string method)
    {
        return new Dictionary<string, object?>
        {
            ["pid"] = pid,
            ["terminated"] = terminated,
            ["method"] = method,
        };
    }
}
