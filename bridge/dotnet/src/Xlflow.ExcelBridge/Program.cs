using System.Text;

namespace Xlflow.ExcelBridge;

public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        using var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: false);
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
        using var stderr = new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false)) { AutoFlush = true };

        return BridgeStartup.Run(
            args,
            stdin,
            stdout,
            stderr,
            Services.ExcelBridgeSupport.RegisterOleMessageFilter,
            Workers.MacroRunWorker.Run,
            (startupArgs, startupStdin, startupStdout, startupStderr) => BridgeHost.Run(startupArgs, startupStdin, startupStdout, startupStderr));
    }
}
