using System.Globalization;
using System.Text;
using System.Text.Json;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class FormWriteContractTests
{
    [Fact]
    public void RollbackFailurePoisonsLiveSessionAndRequiresDiscard()
    {
        var directory = Path.Combine(Path.GetTempPath(), "xlflow-form-rollback-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var metadataPath = Path.Combine(directory, "session.json");
            var workbookPath = Path.Combine(directory, "Book.xlsm");
            File.WriteAllText(metadataPath, JsonSerializer.Serialize(new { hwnd = 123, pid = 456, workbook_path = workbookPath }));
            var args = new FormWriteCommandArguments("apply", workbookPath, "", "", "sidecar", false, "", false, "", true, true, false, true, metadataPath);
            var response = ExcelFormWriteService.FormImportRollbackFailure(BridgeRequest.Unknown, args, true, "owned", "form import failed and rollback failed");
            Assert.Equal("form_import_rollback_failed", response.Error!.Code);
            var metadata = ExcelBridgeSupport.ReadSessionMetadata(metadataPath);
            Assert.NotNull(metadata);
            Assert.True(metadata!.Poisoned);
            Assert.True(metadata.DiscardUnsavedChanges);
            Assert.Throws<SessionPoisonedException>(() => ExcelBridgeSupport.ThrowIfSessionPoisoned(metadataPath, workbookPath));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PersistedClientDimensionsAcceptSignedHimetricBoundary()
    {
        var twips = int.MaxValue * 1440.0 / 2540.0;
        var text = twips.ToString("R", CultureInfo.InvariantCulture);
        var dimensions = ExcelFormInspectionService.ParsePersistedClientDimensions($"Begin VB.UserForm Sample\n ClientWidth = {text}\n ClientHeight = {text}\nEnd\n");
        Assert.Equal(twips / 20.0, dimensions.Width);
    }

    [Fact]
    public void BridgeReadsOnlyTheCanonicalSpecNameAndAcceptsCompilerOwnedFields()
    {
        const string json = """
            {
              "form": {
                "name": "Sample",
                "caption": "Canonical caption",
                "width": 280,
                "height": 180,
                "client_width": 500,
                "client_height": 400,
                "build": { "clientWidth": 500, "clientHeight": 400 }
              },
              "controls": [{ "type": "Image", "picture": { "path": "assets/logo.png" } }]
            }
            """;
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

        Assert.Equal("Sample", ExcelFormWriteService.ReadFormName(encoded));
    }

    [Fact]
    public void BridgeHasNoImperativeDesignerAuthoringFallback()
    {
        var methods = typeof(ExcelFormWriteService).GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Instance);

        Assert.DoesNotContain(methods, method => method.Name is "BuildForm" or "ApplyForm" or "AddDesignerControl");
    }

    [Theory]
    [InlineData(10000, 8000, 500, 400)]
    [InlineData(0, 0, 0, 0)]
    public void PersistedClientDimensionsConvertTwipsToPoints(int clientWidthTwips, int clientHeightTwips, double width, double height)
    {
        var dimensions = ExcelFormInspectionService.ParsePersistedClientDimensions(
            $"VERSION 5.00\r\nBegin VB.UserForm Sample\r\n   ClientHeight = {clientHeightTwips}\r\n   ClientWidth = {clientWidthTwips}\r\nEnd\r\n");

        Assert.Equal(width, dimensions.Width);
        Assert.Equal(height, dimensions.Height);
    }

    [Fact]
    public void PersistedClientDimensionsParseInvariantDecimalTwipsRegardlessOfCurrentCulture()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            var dimensions = ExcelFormInspectionService.ParsePersistedClientDimensions(
                "Begin VB.UserForm Sample\n ClientWidth = 10000.5\n ClientHeight = 8000.25\nEnd\n");

            Assert.Equal(500.025, dimensions.Width);
            Assert.Equal(400.0125, dimensions.Height);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    [Theory]
    [InlineData("-0.5", "8000")]
    [InlineData("10000", "-0.5")]
    [InlineData("NaN", "8000")]
    [InlineData("10000", "Infinity")]
    [InlineData("-Infinity", "8000")]
    [InlineData("1e999", "8000")]
    [InlineData("10000,5", "8000")]
    [InlineData("2147483648", "8000")]
    [InlineData("10000", "2147483648")]
    public void PersistedClientDimensionsRejectNegativeNonFiniteOrNonInvariantValues(string clientWidthTwips, string clientHeightTwips)
    {
        var formText = $"Begin VB.UserForm Sample\n ClientWidth = {clientWidthTwips}\n ClientHeight = {clientHeightTwips}\nEnd\n";

        Assert.Throws<InvalidOperationException>(
            () => ExcelFormInspectionService.ParsePersistedClientDimensions(formText));
    }

    [Fact]
    public void PersistedClientDimensionsIgnoreNestedAttributesAndVbaCodeAfterRootBlock()
    {
        const string formText = """
            VERSION 5.00
            Begin VB.UserForm Sample
               ClientWidth = 10000
               ClientHeight = 8000
               Begin VB.Frame Nested
                  ClientWidth = 5
                  ClientHeight = 6
               End
            End
            Attribute VB_Name = "Sample"
            Private Sub Example()
              ClientWidth = 5
              ClientHeight = 6
            End Sub
            """;

        var dimensions = ExcelFormInspectionService.ParsePersistedClientDimensions(formText);

        Assert.Equal(500.0, dimensions.Width);
        Assert.Equal(400.0, dimensions.Height);
    }

    [Theory]
    [InlineData("ClientWidth = 10000", "ClientWidth")]
    [InlineData("ClientHeight = 8000", "ClientHeight")]
    public void PersistedClientDimensionsRejectDuplicateRootAttributes(string duplicateAttribute, string propertyName)
    {
        var formText = $"Begin VB.UserForm Sample\n ClientWidth = 10000\n ClientHeight = 8000\n {duplicateAttribute}\nEnd\n";

        var exception = Assert.Throws<InvalidOperationException>(
            () => ExcelFormInspectionService.ParsePersistedClientDimensions(formText));

        Assert.Contains(propertyName, exception.Message, StringComparison.Ordinal);
        Assert.Contains("duplicate", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(-1, 8000)]
    [InlineData(10000, -1)]
    public void PersistedClientDimensionsRejectNegativeRootAttributes(int clientWidthTwips, int clientHeightTwips)
    {
        var formText = $"Begin VB.UserForm Sample\n ClientWidth = {clientWidthTwips}\n ClientHeight = {clientHeightTwips}\nEnd\n";

        Assert.Throws<InvalidOperationException>(
            () => ExcelFormInspectionService.ParsePersistedClientDimensions(formText));
    }

    [Fact]
    public void DesignerSnapshotKeepsOuterAndPersistedClientDimensionsSeparate()
    {
        var snapshot = ExcelFormInspectionService.BuildDesignerFormRootSnapshot(
            "Sample", "Caption", width: 540, height: 440, clientWidth: 500, clientHeight: 400, []);

        Assert.Equal(540.0, snapshot["width"]);
        Assert.Equal(440.0, snapshot["height"]);
        Assert.Equal(500.0, snapshot["client_width"]);
        Assert.Equal(400.0, snapshot["client_height"]);
        Assert.Equal("Caption", snapshot["caption"]);
    }
}
