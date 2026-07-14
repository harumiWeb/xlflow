using System.Text.Json;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class RuntimeInjectionHelperTests
{
    [Theory]
    [InlineData("headless", "headless")]
    [InlineData("Headless", "headless")]
    [InlineData("HEADLESS", "headless")]
    [InlineData("ci", "ci")]
    [InlineData("CI", "ci")]
    [InlineData("agent", "agent")]
    [InlineData("Agent", "agent")]
    [InlineData("test", "test")]
    [InlineData("Test", "test")]
    [InlineData("interactive", "interactive")]
    [InlineData("Interactive", "interactive")]
    [InlineData("", "interactive")]
    [InlineData(null, "interactive")]
    [InlineData("unknown", "interactive")]
    [InlineData("  headless  ", "headless")]
    public void NormalizeRuntimeModeReturnsExpectedValues(string? input, string expected)
    {
        Assert.Equal(expected, RuntimeInjectionHelper.NormalizeRuntimeMode(input!));
    }

    [Theory]
    [InlineData("confirm-save", "confirm_save")]
    [InlineData("confirm save", "confirm_save")]
    [InlineData("Confirm Save", "confirm_save")]
    [InlineData("CONFIRM-SAVE", "confirm_save")]
    [InlineData("test123", "test123")]
    [InlineData("a", "a")]
    [InlineData("hello world", "hello_world")]
    [InlineData("a--b", "a_b")]
    [InlineData("a__b", "a_b")]
    [InlineData("  spaces  ", "spaces")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void ConvertToUIResponseIdNormalizesCorrectly(string? input, string expected)
    {
        Assert.Equal(expected, RuntimeInjectionHelper.ConvertToUIResponseId(input!));
    }

    [Fact]
    public void GetUIResponseDefinedNameBuildsMsgBoxName()
    {
        var result = RuntimeInjectionHelper.GetUIResponseDefinedName("msgbox", "confirm-save");
        Assert.Equal("__XLFLOW_UI_MSGBOX_confirm_save__", result);
    }

    [Fact]
    public void GetUIResponseDefinedNameBuildsInputName()
    {
        var result = RuntimeInjectionHelper.GetUIResponseDefinedName("input", "customer-name");
        Assert.Equal("__XLFLOW_UI_INPUT_customer_name__", result);
    }

    [Theory]
    [InlineData("get-open", "source-files", "__XLFLOW_UI_FILEDIALOG_GET_OPEN_source_files__")]
    [InlineData("file-open", "pick-report", "__XLFLOW_UI_FILEDIALOG_FILE_OPEN_pick_report__")]
    [InlineData("save-as", "export-path", "__XLFLOW_UI_FILEDIALOG_SAVE_AS_export_path__")]
    [InlineData("folder", "export-dir", "__XLFLOW_UI_FILEDIALOG_FOLDER_export_dir__")]
    public void GetFileDialogResponseDefinedNameBuildsCorrectName(string kind, string id, string expected)
    {
        Assert.Equal(expected, RuntimeInjectionHelper.GetFileDialogResponseDefinedName(kind, id));
    }

    [Fact]
    public void GetFileDialogResponseDefinedNameRejectsUnsupportedKind()
    {
        Assert.Throws<InvalidOperationException>(() =>
            RuntimeInjectionHelper.GetFileDialogResponseDefinedName("unknown", "test"));
    }

    [Fact]
    public void GetUIResponseDefinedNameRejectsUnsupportedKind()
    {
        Assert.Throws<InvalidOperationException>(() =>
            RuntimeInjectionHelper.GetUIResponseDefinedName("unknown", "test"));
    }

    [Fact]
    public void GetUIResponseDefinedNameRejectsEmptyId()
    {
        Assert.Throws<InvalidOperationException>(() =>
            RuntimeInjectionHelper.GetUIResponseDefinedName("msgbox", "---"));
    }

    [Fact]
    public void BuildDefinedNameRefersToWrapsValueInQuotes()
    {
        Assert.Equal("=\"hello\"", RuntimeInjectionHelper.BuildDefinedNameRefersTo("hello"));
    }

    [Fact]
    public void BuildDefinedNameRefersToEscapesDoubleQuotes()
    {
        Assert.Equal("=\"say \"\"hi\"\"\"", RuntimeInjectionHelper.BuildDefinedNameRefersTo("say \"hi\""));
    }

    [Theory]
    [InlineData("__XLFLOW_MODE__")]
    [InlineData("Book.xlsm!__XLFLOW_MODE__")]
    [InlineData("'Book.xlsm'!__XLFLOW_RUNTIME_VERSION__")]
    [InlineData("__XLFLOW_DEBUG_PIPE__")]
    [InlineData("__XLFLOW_RUN_HELPER__")]
    [InlineData("__XLFLOW_UI_STREAM_HELPER__")]
    [InlineData("__XLFLOW_UI_STREAM_REDACT_INPUT__")]
    [InlineData("__XLFLOW_UI_MSGBOX_confirm_save__")]
    [InlineData("__XLFLOW_UI_INPUT_customer_name__")]
    [InlineData("__XLFLOW_UI_FILEDIALOG_GET_OPEN_source_files__")]
    public void IsTransientRuntimeDefinedNameRecognizesXlflowOwnedMarkers(string name)
    {
        Assert.True(RuntimeInjectionHelper.IsTransientRuntimeDefinedName(name));
    }

    [Theory]
    [InlineData("")]
    [InlineData("CustomerName")]
    [InlineData("XLFLOW_MODE")]
    [InlineData("__XLFLOW_PROJECT_SETTING__")]
    [InlineData("__XLFLOW_UI_CUSTOM__")]
    public void IsTransientRuntimeDefinedNamePreservesUnownedNames(string name)
    {
        Assert.False(RuntimeInjectionHelper.IsTransientRuntimeDefinedName(name));
    }

    [Theory]
    [InlineData("XlflowRun_1234abcd")]
    [InlineData("xlflowuistream_deadbeef")]
    public void IsTemporaryRuntimeComponentNameRecognizesGeneratedHelpers(string name)
    {
        Assert.True(RuntimeInjectionHelper.IsTemporaryRuntimeComponentName(name));
    }

    [Theory]
    [InlineData("XlflowRunner")]
    [InlineData("XlflowRuntime")]
    [InlineData("CustomerModule")]
    public void IsTemporaryRuntimeComponentNamePreservesPersistentModules(string name)
    {
        Assert.False(RuntimeInjectionHelper.IsTemporaryRuntimeComponentName(name));
    }

    [Fact]
    public void RemoveTransientRuntimeDefinedNamesDeletesOnlyXlflowOwnedMarkers()
    {
        var mode = new FakeDefinedName("Book.xlsm!__XLFLOW_MODE__");
        var response = new FakeDefinedName("__XLFLOW_UI_INPUT_customer_name__");
        var customer = new FakeDefinedName("CustomerName");
        var workbook = new FakeWorkbook(mode, response, customer);

        var removed = RuntimeInjectionHelper.RemoveTransientRuntimeDefinedNames(workbook);

        Assert.Equal(2, removed);
        Assert.True(mode.Deleted);
        Assert.True(response.Deleted);
        Assert.False(customer.Deleted);
    }

    [Fact]
    public void WorkbookSaveDoesNotRunWhenTransientMarkerCleanupFails()
    {
        var marker = new FakeDefinedName("__XLFLOW_MODE__") { DeleteFailure = new InvalidOperationException("delete failed") };
        var workbook = new FakeWorkbook(marker);

        var error = Assert.ThrowsAny<Exception>(() => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));

        Assert.Contains("delete failed", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, workbook.SaveCalls);
    }

    [Fact]
    public void WorkbookSaveRemovesTransientMarkersBeforeSaving()
    {
        var marker = new FakeDefinedName("__XLFLOW_MODE__");
        var temporary = new FakeComponent("XlflowRun_1234abcd");
        var persistent = new FakeComponent("XlflowRunner");
        var workbook = new FakeWorkbook([marker], [temporary, persistent]);

        ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");

        Assert.True(marker.Deleted);
        Assert.True(temporary.Removed);
        Assert.False(persistent.Removed);
        Assert.Equal(1, workbook.SaveCalls);
    }

    [Fact]
    public void EnsureDefinedNameRestorationTracksHarnessMarkerWhenModeInjectionWasSkipped()
    {
        var workbook = new FakeWorkbook();
        var state = new RuntimeInjectionHelper.RuntimeInjectionState();

        RuntimeInjectionHelper.EnsureDefinedNameRestoration(workbook, state, "__XLFLOW_RUN_HELPER__");
        RuntimeInjectionHelper.EnsureDefinedNameRestoration(workbook, state, "__XLFLOW_RUN_HELPER__");

        Assert.True(state.RestoreRequired);
        var snapshot = Assert.Single(state.NameSnapshots);
        Assert.Equal("__XLFLOW_RUN_HELPER__", snapshot.Name);
        Assert.False(snapshot.Existed);
    }

    [Fact]
    public void WorkbookSaveDoesNotRequireVBProjectWhenNoTransientMarkersExist()
    {
        var workbook = new FakeWorkbookWithoutVBProject();

        ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save");

        Assert.Equal(1, workbook.SaveCalls);
    }

    [Fact]
    public void WorkbookSaveRequiresVBProjectWhenTransientMarkersExist()
    {
        var workbook = new FakeWorkbookWithoutVBProject(new FakeDefinedName("__XLFLOW_MODE__"));

        Assert.ThrowsAny<Exception>(() => ExcelBridgeSupport.InvokeViaDynamic(workbook, "Save"));

        Assert.Equal(0, workbook.SaveCalls);
    }

    [Fact]
    public void DecodeUIResponsesJsonParsesBase64Json()
    {
        var json = "{\"confirm-save\":\"yes\",\"customer-name\":\"Jane\"}";
        var base64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(json));

        var result = RuntimeInjectionHelper.DecodeUIResponsesJson(base64);

        Assert.Equal(2, result.Count);
        Assert.Equal("yes", result["confirm-save"]);
        Assert.Equal("Jane", result["customer-name"]);
    }

    [Fact]
    public void DecodeUIResponsesJsonReturnsEmptyForNullOrWhitespace()
    {
        Assert.Empty(RuntimeInjectionHelper.DecodeUIResponsesJson(""));
        Assert.Empty(RuntimeInjectionHelper.DecodeUIResponsesJson(null!));
        Assert.Empty(RuntimeInjectionHelper.DecodeUIResponsesJson("   "));
    }

    [Fact]
    public void DecodeUIResponsesJsonReturnsEmptyForInvalidBase64()
    {
        Assert.Empty(RuntimeInjectionHelper.DecodeUIResponsesJson("not-valid-base64!"));
    }

    [Fact]
    public void DecodeFileDialogResponsesJsonParsesArray()
    {
        var json = """[{"kind":"file-open","dialog_id":"pick-report","cancelled":false,"values":["C:\\a.txt","C:\\b.txt"]}]""";
        var base64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(json));

        var result = RuntimeInjectionHelper.DecodeFileDialogResponsesJson(base64);

        Assert.Single(result);
        Assert.Equal("file-open", result[0].Kind);
        Assert.Equal("pick-report", result[0].DialogId);
        Assert.False(result[0].Cancelled);
        Assert.Equal(2, result[0].Values.Count);
        Assert.Equal("C:\\a.txt", result[0].Values[0]);
        Assert.Equal("C:\\b.txt", result[0].Values[1]);
    }

    [Fact]
    public void DecodeFileDialogResponsesJsonHandlesCancelled()
    {
        var json = """[{"kind":"folder","dialog_id":"export-dir","cancelled":true,"values":[]}]""";
        var base64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(json));

        var result = RuntimeInjectionHelper.DecodeFileDialogResponsesJson(base64);

        Assert.Single(result);
        Assert.True(result[0].Cancelled);
    }

    [Fact]
    public void DecodeFileDialogResponsesJsonReturnsEmptyForNullOrWhitespace()
    {
        Assert.Empty(RuntimeInjectionHelper.DecodeFileDialogResponsesJson(""));
        Assert.Empty(RuntimeInjectionHelper.DecodeFileDialogResponsesJson(null!));
    }

    [Fact]
    public void ConvertToFileDialogMarkerValueReturnsCancelForCancelled()
    {
        var response = new RuntimeInjectionHelper.FileDialogResponse("folder", "test", true, []);
        Assert.Equal("@cancel", RuntimeInjectionHelper.ConvertToFileDialogMarkerValue(response));
    }

    [Fact]
    public void ConvertToFileDialogMarkerValueJoinsValuesWithNewline()
    {
        var response = new RuntimeInjectionHelper.FileDialogResponse("file-open", "test", false, ["C:\\a.txt", "C:\\b.txt"]);
        Assert.Equal("C:\\a.txt\nC:\\b.txt", RuntimeInjectionHelper.ConvertToFileDialogMarkerValue(response));
    }

    [Fact]
    public void ConvertToFileDialogMarkerValueReturnsSingleValueWithoutNewline()
    {
        var response = new RuntimeInjectionHelper.FileDialogResponse("file-open", "test", false, ["C:\\a.txt"]);
        Assert.Equal("C:\\a.txt", RuntimeInjectionHelper.ConvertToFileDialogMarkerValue(response));
    }

    [Fact]
    public void BuildUIStreamModuleCodeContainsPipeName()
    {
        var code = RuntimeInjectionHelper.BuildUIStreamModuleCode(@"\\.\pipe\xlflow-ui-test");

        Assert.Contains(@"\\.\pipe\xlflow-ui-test", code);
        Assert.Contains("Option Explicit", code);
        Assert.Contains("EmitEvent", code);
        Assert.Contains("CreateFileW", code);
        Assert.Contains("WriteFile", code);
        Assert.Contains("CloseHandle", code);
    }

    [Fact]
    public void BuildUIStreamModuleCodeEscapesQuotesInPipeName()
    {
        var code = RuntimeInjectionHelper.BuildUIStreamModuleCode("test\"pipe");
        Assert.Contains("test\"\"pipe", code);
    }

    public sealed class FakeWorkbook
    {
        public FakeWorkbook(params FakeDefinedName[] names) : this(names, [])
        {
        }

        public FakeWorkbook(FakeDefinedName[] names, FakeComponent[] components)
        {
            Names = new FakeNames(names);
            VBProject = new FakeVBProject(components);
        }

        public FakeNames Names { get; }
        public FakeVBProject VBProject { get; }
        public int SaveCalls { get; private set; }

        public void Save()
        {
            SaveCalls++;
        }
    }

    public sealed class FakeNames(params FakeDefinedName[] names)
    {
        private readonly FakeDefinedName[] items = names;
        public int Count => items.Length;
        public FakeDefinedName Item(int index) => items[index - 1];
    }

    public sealed class FakeDefinedName(string name)
    {
        public string Name { get; } = name;
        public bool Deleted { get; private set; }
        public Exception? DeleteFailure { get; init; }

        public void Delete()
        {
            if (DeleteFailure is not null)
            {
                throw DeleteFailure;
            }
            Deleted = true;
        }
    }

    public sealed class FakeVBProject(FakeComponent[] components)
    {
        public FakeComponents VBComponents { get; } = new(components);
    }

    public sealed class FakeComponents(FakeComponent[] components)
    {
        private readonly FakeComponent[] items = components;
        public int Count => items.Length;
        public FakeComponent Item(int index) => items[index - 1];

        public object? Remove(object component)
        {
            ((FakeComponent)component).Removed = true;
            return null;
        }
    }

    public sealed class FakeComponent(string name)
    {
        public string Name { get; } = name;
        public bool Removed { get; set; }
    }

    public sealed class FakeWorkbookWithoutVBProject(params FakeDefinedName[] names)
    {
        public FakeNames Names { get; } = new(names);
        public int SaveCalls { get; private set; }

        public void Save()
        {
            SaveCalls++;
        }
    }
}
