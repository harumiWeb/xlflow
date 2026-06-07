namespace Xlflow.ExcelBridge;

public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        using var messageFilter = Services.ExcelBridgeSupport.RegisterOleMessageFilter();
        if (args.Contains("--run-worker", StringComparer.OrdinalIgnoreCase))
        {
            return Workers.MacroRunWorker.Run();
        }
        return BridgeHost.Run(args, Console.In, Console.Out, Console.Error);
    }
}
