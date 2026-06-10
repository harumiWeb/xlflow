using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge.Tests;

public sealed class BridgeStartupTests
{
    [Fact]
    public void NoArgumentsPrintsExplanationAndExitsZero()
    {
        using var stdin = TextReader.Null;
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var registerCalls = 0;
        var hostCalls = 0;

        var code = BridgeStartup.Run([], stdin, stdout, stderr, () => new CallbackDisposable(() => registerCalls++), () => 99, (_, _, _, _) =>
        {
            hostCalls++;
            return 98;
        });

        Assert.Equal(0, code);
        Assert.Contains("internal helper executable", stdout.ToString());
        Assert.Contains("Please run it through `xlflow`.", stdout.ToString());
        Assert.Equal(string.Empty, stderr.ToString());
        Assert.Equal(0, registerCalls);
        Assert.Equal(0, hostCalls);
    }

    [Theory]
    [InlineData("--help")]
    [InlineData("-h")]
    public void HelpFlagsPrintUsageAndExitZero(string arg)
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeStartup.Run([arg], TextReader.Null, stdout, stderr, NoopDisposable, () => 99, (_, _, _, _) => 98);

        Assert.Equal(0, code);
        Assert.Contains("Usage:", stdout.ToString());
        Assert.Equal(string.Empty, stderr.ToString());
    }

    [Fact]
    public void VersionPrintsHumanReadableVersionAndExitsZero()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var code = BridgeStartup.Run(["--version"], TextReader.Null, stdout, stderr, NoopDisposable, () => 99, (_, _, _, _) => 98);

        Assert.Equal(0, code);
        Assert.Contains("xlflow-excel-bridge", stdout.ToString());
        Assert.Contains(BridgeInfo.Current().Version, stdout.ToString());
        Assert.Equal(string.Empty, stderr.ToString());
    }

    [Fact]
    public void InvalidArgumentsFailFastWithoutInitializingRuntime()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var registerCalls = 0;

        var code = BridgeStartup.Run(["--bogus"], TextReader.Null, stdout, stderr, () => new CallbackDisposable(() => registerCalls++), () => 99, (_, _, _, _) => 98);

        Assert.Equal(2, code);
        Assert.Equal(string.Empty, stdout.ToString());
        Assert.Contains("Invalid arguments", stderr.ToString());
        Assert.Equal(0, registerCalls);
    }

    [Fact]
    public void InternalRunFlagStartsBridgeHost()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var registerCalls = 0;
        var hostCalls = 0;
        string[]? forwardedArgs = null;

        var code = BridgeStartup.Run([BridgeStartup.InternalRunFlag], TextReader.Null, stdout, stderr, () => new CallbackDisposable(() => registerCalls++), () => 99, (args, _, _, _) =>
        {
            hostCalls++;
            forwardedArgs = args;
            return 7;
        });

        Assert.Equal(7, code);
        Assert.Equal(1, registerCalls);
        Assert.Equal(1, hostCalls);
        Assert.NotNull(forwardedArgs);
        Assert.Empty(forwardedArgs!);
    }

    [Fact]
    public void InternalWorkerModeRequiresInternalFlag()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var workerCalls = 0;

        var code = BridgeStartup.Run([BridgeStartup.InternalRunFlag, "--run-worker"], TextReader.Null, stdout, stderr, NoopDisposable, () =>
        {
            workerCalls++;
            return 5;
        }, (_, _, _, _) => 98);

        Assert.Equal(5, code);
        Assert.Equal(1, workerCalls);
    }

    private static IDisposable NoopDisposable()
    {
        return new CallbackDisposable(() => { });
    }

    private sealed class CallbackDisposable(Action callback) : IDisposable
    {
        public void Dispose()
        {
            callback();
        }
    }
}
