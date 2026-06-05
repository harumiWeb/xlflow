using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelInspectionServiceTests
{
    [Fact]
    public void NormalizeCellValueKeepsPrimitiveScalars()
    {
        Assert.Equal(42d, ExcelBridgeSupport.NormalizeCellValue(42d));
        Assert.Equal(true, ExcelBridgeSupport.NormalizeCellValue(true));
        Assert.Equal("abc", ExcelBridgeSupport.NormalizeCellValue("abc"));
        Assert.Null(ExcelBridgeSupport.NormalizeCellValue(null));
    }
}
