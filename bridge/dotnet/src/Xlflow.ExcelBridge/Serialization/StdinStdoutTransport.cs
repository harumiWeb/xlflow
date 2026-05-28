using System.Text.Json;

namespace Xlflow.ExcelBridge.Serialization;

public static class StdinStdoutTransport
{
    public static T? Read<T>(TextReader reader)
    {
        var input = reader.ReadToEnd();
        if (string.IsNullOrWhiteSpace(input))
        {
            return default;
        }

        return JsonSerializer.Deserialize<T>(input, JsonOptions.Default);
    }

    public static void Write<T>(TextWriter writer, T value)
    {
        writer.WriteLine(JsonSerializer.Serialize(value, JsonOptions.Default));
    }
}
