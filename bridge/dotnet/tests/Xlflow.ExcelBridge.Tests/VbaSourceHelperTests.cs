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
            var exported = "Attribute VB_Name = \"Module1\"\nSub Hello()\n    MsgBox \"日本語テスト\"\nEnd Sub\n";
            File.WriteAllText(tempFile, exported, VbaSourceHelper.GetVbaInteropEncoding());

            var content = VbaSourceHelper.ReadExportedTextAsUtf8(tempFile);

            Assert.Equal(exported.Replace("\n", "\r\n"), content);
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
    public void GetVbaImportEncoding_ReturnsInteropEncoding()
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
}
