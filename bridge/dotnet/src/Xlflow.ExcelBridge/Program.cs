using System.Text;

namespace Xlflow.ExcelBridge;

public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        using var messageFilter = Services.ExcelBridgeSupport.RegisterOleMessageFilter();
        using var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: false);
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        using var stderr = new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false)) { AutoFlush = true };

        if (args.Contains("--run-worker", StringComparer.OrdinalIgnoreCase))
        {
            return Workers.MacroRunWorker.Run();
        }

        return BridgeHost.Run(args, stdin, stdout, stderr);
    }
}
