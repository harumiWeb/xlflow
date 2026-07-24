namespace Xlflow.ExcelBridge.Services;

// The plan is produced by internal/build in the Go frontend.  The bridge must
// consume this explicit payload and must never rediscover source roots itself.
public sealed record BuildCommandArguments(
    string BaseWorkbookPath,
    string TemporaryDirectory,
    string PlanJson64,
    string CodeSource,
    bool Visible);
