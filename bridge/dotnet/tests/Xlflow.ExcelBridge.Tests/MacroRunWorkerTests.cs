using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Tests;

public sealed class MacroRunWorkerTests
{
    [Fact]
    public void CreateStartInfoUsesInternalWorkerMode()
    {
        var startInfo = MacroRunWorkerProcess.CreateStartInfo("request.json", "result.json");

        Assert.False(startInfo.UseShellExecute);
        Assert.Contains("--run-worker", startInfo.ArgumentList);
        Assert.Equal("request.json", startInfo.Environment["XLFLOW_WORKER_REQUEST_PATH"]);
        Assert.Equal("result.json", startInfo.Environment["XLFLOW_WORKER_RESULT_PATH"]);
    }
}
