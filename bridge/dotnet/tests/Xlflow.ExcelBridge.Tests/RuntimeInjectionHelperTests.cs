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
}
