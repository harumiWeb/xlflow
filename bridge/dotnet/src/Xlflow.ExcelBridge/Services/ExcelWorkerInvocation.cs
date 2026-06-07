using System.Diagnostics;
using Xlflow.ExcelBridge.Windows;
using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Services;

internal sealed record WorkerInvocationResult(
    MacroRunWorkerResult? Result,
    DialogSnapshot? Dialog,
    DialogSnapshot[] Dialogs,
    VbeSelectionCapture LocationCapture,
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
        CancellationToken cancellationToken,
        IVbeSelectionLocator? selectionLocator = null)
    {
        using var worker = MacroRunWorkerProcess.Start(workerRequest);
        var watcher = new DialogWatcher();
        var actionPolicy = suppressModalErrors
            ? selectionLocator is not null
                ? DialogActionPolicy.SuppressVbaErrorWithRuntimeDebug
                : DialogActionPolicy.SuppressVbaError
            : DialogActionPolicy.ObserveOnly;
        var watchRequest = new DialogWatchRequest(
            workerRequest.ExcelProcessId,
            excelHwnd,
            dialogKind,
            actionPolicy,
            timeout,
            TimeSpan.FromMilliseconds(50));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var locationCaptures = new List<VbeSelectionCapture>();
        void CaptureBefore(DialogKind kind)
        {
            if ((kind is DialogKind.Compile or DialogKind.Runtime) && selectionLocator is not null)
            {
                locationCaptures.Add(selectionLocator.Capture("before_dialog_action"));
            }
        }

        void CaptureAfter(DialogKind kind)
        {
            if ((kind is DialogKind.Compile or DialogKind.Runtime) &&
                selectionLocator is not null &&
                !locationCaptures.Any(capture => capture.HasReliableLocation))
            {
                locationCaptures.Add(selectionLocator.Capture("after_dialog_action"));
            }
            if (kind == DialogKind.Runtime)
            {
                selectionLocator?.ResetBreakMode();
            }
        }

        var watcherTask = Task.Run(() => watcher.WaitForDialog(watchRequest, CaptureBefore, CaptureAfter, linked.Token), linked.Token);
        var stopwatch = Stopwatch.StartNew();

        while (stopwatch.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (watcherTask.IsCompletedSuccessfully && watcherTask.Result is not null)
            {
                worker.Stop();
                linked.Cancel();
                return new WorkerInvocationResult(null, watcherTask.Result, [watcherTask.Result], MergeCaptures(locationCaptures), false, worker.ProcessId);
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
                    linked.Token,
                    CaptureBefore,
                    CaptureAfter);
                linked.Cancel();
                if (postDialog is not null)
                {
                    return new WorkerInvocationResult(result, postDialog, [postDialog], MergeCaptures(locationCaptures), false, worker.ProcessId);
                }
                return new WorkerInvocationResult(result, null, [], MergeCaptures(locationCaptures), false, worker.ProcessId);
            }
            Thread.Sleep(25);
        }

        worker.Stop();
        linked.Cancel();
        var dialogs = watcher.CaptureCurrentDialogs(watchRequest, includeUia: false).ToArray();
        return new WorkerInvocationResult(null, dialogs.FirstOrDefault(), dialogs, MergeCaptures(locationCaptures), true, worker.ProcessId);
    }

    internal static DialogSnapshot? WaitForPostWorkerDialog(
        DialogWatcher watcher,
        Task<DialogSnapshot?> watcherTask,
        DialogWatchRequest watchRequest,
        string operation,
        MacroRunWorkerResult? result,
        CancellationToken cancellationToken,
        Action<DialogKind>? beforeAction = null,
        Action<DialogKind>? afterAction = null)
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
            var dialog = watcher.TryCaptureCurrentDialog(watchRequest, includeUia: true, executeAction: true, beforeAction, afterAction);
            if (dialog is not null)
            {
                return dialog;
            }
            Thread.Sleep(25);
        }
        return null;
    }

    private static VbeSelectionCapture MergeCaptures(List<VbeSelectionCapture> captures)
    {
        if (captures.Count == 0)
        {
            return VbeSelectionCapture.Empty;
        }

        var location = captures
            .Select(capture => capture.Location)
            .Where(location => location is not null)
            .OrderByDescending(location => VbeSelectionScorer.Score(location!))
            .FirstOrDefault();
        return new VbeSelectionCapture(location, captures.SelectMany(capture => capture.Attempts).ToArray());
    }
}
