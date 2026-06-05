using System.Diagnostics;
using Xlflow.ExcelBridge.Windows;
using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Services;

internal sealed record WorkerInvocationResult(
    MacroRunWorkerResult? Result,
    DialogSnapshot? Dialog,
    DialogSnapshot[] Dialogs,
    bool TimedOut,
    int WorkerProcessId);

internal static class ExcelWorkerInvocation
{
    internal static WorkerInvocationResult InvokeWithWorker(
        MacroRunWorkerRequest workerRequest,
        long excelHwnd,
        DialogKind dialogKind,
        bool suppressModalErrors,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        using var worker = MacroRunWorkerProcess.Start(workerRequest);
        var watcher = new DialogWatcher();
        var watchRequest = new DialogWatchRequest(
            workerRequest.ExcelProcessId,
            excelHwnd,
            dialogKind,
            suppressModalErrors ? DialogActionPolicy.SuppressVbaError : DialogActionPolicy.ObserveOnly,
            timeout,
            TimeSpan.FromMilliseconds(50));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var watcherTask = Task.Run(() => watcher.WaitForDialog(watchRequest, linked.Token), linked.Token);
        var stopwatch = Stopwatch.StartNew();

        while (stopwatch.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
            {
                worker.Stop();
                linked.Cancel();
                return new WorkerInvocationResult(null, watcherTask.Result, [watcherTask.Result], false, worker.ProcessId);
            }
            if (worker.HasExited)
            {
                var result = worker.WaitForResult(TimeSpan.FromSeconds(1));
                var postDialog = WaitForPostWorkerDialog(
                    watcher,
                    watcherTask,
                    watchRequest,
                    workerRequest.Operation,
                    result,
                    linked.Token);
                linked.Cancel();
                if (postDialog is not null)
                {
                    return new WorkerInvocationResult(result, postDialog, [postDialog], false, worker.ProcessId);
                }
                return new WorkerInvocationResult(result, null, [], false, worker.ProcessId);
            }
            Thread.Sleep(25);
        }

        worker.Stop();
        linked.Cancel();
        var dialogs = watcher.CaptureCurrentDialogs(watchRequest, includeUia: false).ToArray();
        return new WorkerInvocationResult(null, dialogs.FirstOrDefault(), dialogs, true, worker.ProcessId);
    }

    internal static DialogSnapshot? WaitForPostWorkerDialog(
        DialogWatcher watcher,
        Task<DialogSnapshot?> watcherTask,
        DialogWatchRequest watchRequest,
        string operation,
        MacroRunWorkerResult? result,
        CancellationToken cancellationToken)
    {
        var shouldWait =
            string.Equals(operation, "compile", StringComparison.OrdinalIgnoreCase) ||
            result is null ||
            !result.Ok ||
            result.Error is not null;
        if (!shouldWait)
        {
            return null;
        }

        if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
        {
            return watcherTask.Result;
        }

        var deadline = DateTime.UtcNow.AddMilliseconds(900);
        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
            {
                return watcherTask.Result;
            }
            var dialog = watcher.TryCaptureCurrentDialog(watchRequest, includeUia: true, executeAction: true);
            if (dialog is not null)
            {
                return dialog;
            }
            Thread.Sleep(25);
        }
        return null;
    }
}
