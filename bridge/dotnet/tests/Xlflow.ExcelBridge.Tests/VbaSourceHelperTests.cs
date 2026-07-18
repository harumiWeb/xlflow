using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class VbaSourceHelperTests
{
    [Fact]
    public void ReadExportedTextAsUtf8_DecodesCp932Source()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"xlflow-vba-export-{Guid.NewGuid():N}.bas");
        try
        {
            var exported = "Attribute VB_Name = \"Module1\"\r\nSub Hello()\r\n    MsgBox \"日本語テスト\"\r\nEnd Sub\r\n";
            File.WriteAllText(tempFile, exported, VbaSourceHelper.GetVbaInteropEncoding());

            var content = VbaSourceHelper.ReadExportedTextAsUtf8(tempFile);

            Assert.Equal(exported, content);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public void GetVbaImportEncoding_HasSameCodePageAsInteropEncoding()
    {
        Assert.Equal(VbaSourceHelper.GetVbaInteropEncoding().CodePage, VbaSourceHelper.GetVbaImportEncoding().CodePage);
    }

    [Fact]
    public void FindDuplicateModuleNames_IgnoresUserFormCodeSidecars()
    {
        var files = new List<DiscoveredSourceFile>
        {
            new()
            {
                Kind = "form",
                RelativePath = "CustomerForm.frm",
                Extension = ".frm",
                ModuleName = "CustomerForm",
            },
            new()
            {
                Kind = "form_code",
                RelativePath = "code/CustomerForm.bas",
                Extension = ".bas",
                ModuleName = "CustomerForm",
            },
        };

        var duplicates = VbaSourceHelper.FindDuplicateModuleNames(files);

        Assert.Empty(duplicates);
    }

    [Fact]
    public void ErlLineNumbers_AddsPhysicalSourceLinesWithoutReformatting()
    {
        var source = "Option Explicit\r\n\r\nPublic Sub Demo()\r\n    Dim total As Long\r\n    total = 1 _\r\n        + 2\r\n    Debug.Print total\r\nEnd Sub\r\n";

        var applied = ErlLineNumberTransformer.TryAdd(source, out var numbered, out var issue);

        Assert.True(applied);
        Assert.Null(issue);
        Assert.Equal("Option Explicit\r\n\r\nPublic Sub Demo()\r\n    Dim total As Long\r\n5     total = 1 _\r\n        + 2\r\n7     Debug.Print total\r\nEnd Sub\r\n", numbered);
        Assert.True(ErlLineNumberTransformer.TryRemove(numbered, out var restored, out issue));
        Assert.Null(issue);
        Assert.Equal(source, restored);

        var exported = numbered.Replace("5     total", "5      total", StringComparison.Ordinal).Replace("7     Debug", "7      Debug", StringComparison.Ordinal);
        Assert.True(ErlLineNumberTransformer.TryRemove(exported, out restored, out issue, excelExported: true));
        Assert.Equal(source, restored);
    }

    [Fact]
    public void ErlLineNumbers_UseModuleWidthAndRejectUnsafeLegacyNumbers()
    {
        var source = "' one\n' two\n' three\n' four\n' five\n' six\n' seven\n' eight\nPublic Sub Demo()\n    Debug.Print \"ok\"\nEnd Sub\n";
        Assert.True(ErlLineNumberTransformer.TryAdd(source, out var numbered, out var issue));
        Assert.Null(issue);
        Assert.Contains("10     Debug.Print", numbered);

        var unsafeSource = "Public Sub Demo()\n10  Debug.Print \"legacy\"\nEnd Sub\n";
        Assert.False(ErlLineNumberTransformer.TryAdd(unsafeSource, out _, out issue));
        Assert.NotNull(issue);
        Assert.Equal(2, issue!.Line);

        var numericTarget = "Public Sub Demo()\n    GoTo 10\nEnd Sub\n";
        Assert.False(ErlLineNumberTransformer.TryAdd(numericTarget, out _, out issue));
        Assert.NotNull(issue);
        Assert.Equal(2, issue!.Line);

        var unsafeExport = "Public Sub Demo()\n20  Debug.Print \"legacy\"\nEnd Sub\n";
        Assert.False(ErlLineNumberTransformer.TryRemove(unsafeExport, out _, out issue, excelExported: true));
        Assert.NotNull(issue);
        Assert.Equal(2, issue!.Line);

        var errorReset = "Public Sub Demo()\n    On Error GoTo 0\nEnd Sub\n";
        Assert.True(ErlLineNumberTransformer.TryAdd(errorReset, out _, out issue));
        Assert.Null(issue);
    }

    [Fact]
    public void ErlLineNumbers_SkipStructuralLinesAndProcedureDeclarations()
    {
        var source = "Public Sub Demo()\n    If True Then\n        Debug.Print \"yes\"\n    Else\n        Debug.Print \"no\"\n    End If\nEnd Sub\n";

        Assert.True(ErlLineNumberTransformer.TryAdd(source, out var numbered, out var issue));
        Assert.Null(issue);
        Assert.DoesNotContain("2  ", numbered);
        Assert.DoesNotContain("4  ", numbered);
        Assert.DoesNotContain("6  ", numbered);
        Assert.Contains("3 " + "        Debug.Print", numbered);
        Assert.Contains("5 " + "        Debug.Print", numbered);
    }

    [Fact]
    public void PrepareSourceForImport_NumbersAfterFolderAnnotationIsUpdated()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-line-number-annotation-" + Guid.NewGuid().ToString("N"));
        try
        {
            var sourcePath = Path.Combine(root, "Nested", "Main.bas");
            var destinationPath = Path.Combine(root, "import", "Main.bas");
            var source = "Attribute VB_Name = \"Main\"\r\nPublic Sub Run()\r\n    Debug.Print \"ok\"\r\nEnd Sub\r\n";
            Directory.CreateDirectory(Path.GetDirectoryName(sourcePath)!);
            File.WriteAllText(sourcePath, source);

            VbaSourceHelper.PrepareSourceForImport(sourcePath, destinationPath, root, "update", lineNumbersEnabled: true);

            var prepared = File.ReadAllText(destinationPath, VbaSourceHelper.GetVbaInteropEncoding());
            Assert.True(ErlLineNumberTransformer.TryRemove(prepared, out var restored, out var issue));
            Assert.Null(issue);
            Assert.Equal(
                VbaSourceHelper.UpdateFolderAnnotationText(source, "update", "Nested"),
                restored);
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }

    [Fact]
    public void Fingerprint_ChangesWhenErlInstrumentationIsToggled()
    {
        var root = Path.Combine(Path.GetTempPath(), "xlflow-fingerprint-" + Guid.NewGuid().ToString("N"));
        try
        {
            var modules = Path.Combine(root, "modules");
            Directory.CreateDirectory(modules);
            File.WriteAllText(Path.Combine(modules, "Main.bas"), "Attribute VB_Name = \"Main\"\r\n");

            var disabled = VbaSourceHelper.ComputeFingerprint("Book.xlsm", modules, "", "", "", "", false);
            var enabled = VbaSourceHelper.ComputeFingerprint("Book.xlsm", modules, "", "", "", "", true);
            var statePath = Path.Combine(root, "state", "push.json");
            VbaSourceHelper.WriteFingerprintState(disabled, statePath);

            Assert.NotEqual(disabled.LineNumbersEnabled, enabled.LineNumbersEnabled);
            Assert.True(VbaSourceHelper.FingerprintMatchesState(disabled, statePath));
            Assert.False(VbaSourceHelper.FingerprintMatchesState(enabled, statePath));
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }
}
