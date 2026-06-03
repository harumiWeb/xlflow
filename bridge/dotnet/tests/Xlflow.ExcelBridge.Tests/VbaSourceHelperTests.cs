using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class VbaSourceHelperTests
{
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
