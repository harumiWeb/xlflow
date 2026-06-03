using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Tests;

public sealed class MacroRunWorkerTests
{
    [Fact]
    public void CreateStartInfoUsesInternalWorkerMode()
    {
        var startInfo = MacroRunWorkerProcess.CreateStartInfo();

        Assert.False(startInfo.UseShellExecute);
        Assert.True(startInfo.RedirectStandardInput);
        Assert.True(startInfo.RedirectStandardOutput);
        Assert.Contains("--run-worker", startInfo.ArgumentList);
    }
}
