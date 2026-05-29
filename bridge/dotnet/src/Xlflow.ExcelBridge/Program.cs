namespace Xlflow.ExcelBridge;

public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        return BridgeHost.Run(args, Console.In, Console.Out, Console.Error);
    }
}
