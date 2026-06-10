using Xlflow.ExcelBridge.Contract;

namespace Xlflow.ExcelBridge;

internal static class BridgeStartup
{
    internal const string InternalRunFlag = "--bridge-internal-run";
    private static readonly string[] HelpFlags = ["--help", "-h", "/?"];

    internal static int Run(
        string[] args,
        TextReader stdin,
        TextWriter stdout,
        TextWriter stderr,
        Func<IDisposable> registerMessageFilter,
        Func<int> runWorker,
        Func<string[], TextReader, TextWriter, TextWriter, int> runBridgeHost)
    {
        if (HasAny(args, HelpFlags))
        {
            WriteHelp(stdout);
            return 0;
        }

        if (Has(args, "--version"))
        {
            stdout.WriteLine($"{BridgeInfo.Current().Name} {BridgeInfo.Current().Version}");
            return 0;
        }

        if (args.Length == 0)
        {
            stdout.WriteLine("xlflow-excel-bridge is an internal helper executable for xlflow.");
            stdout.WriteLine("Please run it through `xlflow`.");
            return 0;
        }

        if (Has(args, "--version-json") || Has(args, "--capabilities-json"))
        {
            return runBridgeHost(args, stdin, stdout, stderr);
        }

        if (!Has(args, InternalRunFlag))
        {
            stderr.WriteLine("Invalid arguments. Use --help for usage.");
            return 2;
        }

        var runtimeArgs = args.Where(arg => !string.Equals(arg, InternalRunFlag, StringComparison.OrdinalIgnoreCase)).ToArray();

        using var messageFilter = registerMessageFilter();
        if (Has(runtimeArgs, "--run-worker"))
        {
            return runWorker();
        }

        return runBridgeHost(runtimeArgs, stdin, stdout, stderr);
    }

    private static bool Has(string[] args, string value)
    {
        return args.Contains(value, StringComparer.OrdinalIgnoreCase);
    }

    private static bool HasAny(string[] args, IEnumerable<string> values)
    {
        foreach (var value in values)
        {
            if (Has(args, value))
            {
                return true;
            }
        }

        return false;
    }

    private static void WriteHelp(TextWriter stdout)
    {
        stdout.WriteLine("xlflow-excel-bridge is an internal helper executable for xlflow.");
        stdout.WriteLine();
        stdout.WriteLine("Usage:");
        stdout.WriteLine("  xlflow-excel-bridge --help");
        stdout.WriteLine("  xlflow-excel-bridge --version");
        stdout.WriteLine("  xlflow-excel-bridge --version-json");
        stdout.WriteLine("  xlflow-excel-bridge --capabilities-json");
        stdout.WriteLine();
        stdout.WriteLine("Run bridge operations through xlflow.");
    }
}
