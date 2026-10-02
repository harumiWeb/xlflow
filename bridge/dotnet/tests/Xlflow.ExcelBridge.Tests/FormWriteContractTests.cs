using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class FormWriteContractTests
{
    [Theory]
    [InlineData("true", true)]
    [InlineData("false", false)]
    [InlineData("42", 42L)]
    [InlineData("\"Alice\"", "Alice")]
    public void FormValuesAreConvertedToComScalars(string json, object expected)
    {
        using var document = JsonDocument.Parse(json);
        var method = typeof(ExcelFormWriteService).GetMethod("ConvertFormValue", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        Assert.Equal(expected, method!.Invoke(null, [document.RootElement]));
    }

    [Fact]
    public void SharedContractSnapshotMatchesBridgeProgIdsAndDtos()
    {
        var snapshotPath = Path.Combine(AppContext.BaseDirectory, "contract-snapshot.json");
        var snapshot = JsonSerializer.Deserialize<ContractSnapshot>(File.ReadAllText(snapshotPath), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("contract snapshot was empty");

        var mapField = typeof(ExcelFormWriteService).GetField("TypeProgIdMap", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(mapField);
        var map = Assert.IsType<Dictionary<string, string>>(mapField!.GetValue(null));
        Assert.Equal(snapshot.BuiltInControls.Count, map.Count);
        foreach (var control in snapshot.BuiltInControls)
        {
            Assert.True(map.TryGetValue(control.Type, out var progId), $"bridge is missing {control.Type}");
            Assert.Equal(control.ProgId, progId);
        }

        AssertJsonProperties("FormWriteFormSpec", snapshot.Bridge.FormFields);
        AssertJsonProperties("FormWriteBuildFormSpec", snapshot.Bridge.FormBuildFields);
        AssertJsonProperties("FormWriteObservedFormSpec", snapshot.Bridge.ObservedFormFields);
        AssertJsonProperties("FormWriteControlSpec", snapshot.Bridge.ControlFields);
    }

    [Theory]
    [InlineData("clientWidth")]
    [InlineData("clientHeight")]
    public void CompilerClientDimensionsAreRejectedBeforeDesignerAuthoring(string field)
    {
        var method = typeof(ExcelFormWriteService).GetMethod("DecodeSpec", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        var json = JsonSerializer.Serialize(new
        {
            form = new { name = "Sample", build = new Dictionary<string, double> { [field] = 240 } },
        });
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

        var invocation = Assert.Throws<TargetInvocationException>(() => method!.Invoke(null, [encoded]));

        var error = Assert.IsType<InvalidOperationException>(invocation.InnerException);
        Assert.Contains("supported only by pure-Go generation", error.Message);
    }

    [Fact]
    public void ObservedClientDimensionsArePreservedWithoutBecomingBuildIntent()
    {
        var method = typeof(ExcelFormWriteService).GetMethod("DecodeSpec", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        const string json = """{"form":{"name":"Sample","observed":{"insideWidth":238,"insideHeight":175,"clientWidth":240,"clientHeight":180}}}""";
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

        var spec = method!.Invoke(null, [encoded]);

        Assert.NotNull(spec);
        using var document = JsonDocument.Parse(JsonSerializer.Serialize(spec, spec!.GetType()));
        var form = document.RootElement.GetProperty("form");
        Assert.Equal(JsonValueKind.Null, form.GetProperty("build").ValueKind);
        var observed = form.GetProperty("observed");
        Assert.Equal(238, observed.GetProperty("insideWidth").GetDouble());
        Assert.Equal(175, observed.GetProperty("insideHeight").GetDouble());
        Assert.Equal(240, observed.GetProperty("clientWidth").GetDouble());
        Assert.Equal(180, observed.GetProperty("clientHeight").GetDouble());
    }

    private static void AssertJsonProperties(string nestedTypeName, IReadOnlyCollection<string> expected)
    {
        var type = typeof(ExcelFormWriteService).GetNestedType(nestedTypeName, BindingFlags.NonPublic);
        Assert.NotNull(type);
        var actual = type!.GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Select(property => property.GetCustomAttribute<JsonPropertyNameAttribute>()?.Name)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Cast<string>()
            .ToHashSet(StringComparer.Ordinal);
        foreach (var name in expected)
        {
            Assert.Contains(name, actual);
        }
    }

    private sealed class ContractSnapshot
    {
        public List<BuiltInControl> BuiltInControls { get; init; } = [];
        public BridgeContract Bridge { get; init; } = new();
    }

    private sealed class BuiltInControl
    {
        public string Type { get; init; } = "";
        public string ProgId { get; init; } = "";
    }

    private sealed class BridgeContract
    {
        public List<string> FormFields { get; init; } = [];
        public List<string> FormBuildFields { get; init; } = [];
        public List<string> ObservedFormFields { get; init; } = [];
        public List<string> ControlFields { get; init; } = [];
    }
}
