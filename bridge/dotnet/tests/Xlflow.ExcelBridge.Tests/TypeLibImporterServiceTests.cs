using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class TypeLibImporterServiceTests
{
    [Fact]
    public void SelectProgIDsForTypeLibMapsRegisteredProgIDsToCoClassTypes()
    {
        var classID = Guid.Parse("{00024500-0000-0000-C000-000000000046}");
        var otherClassID = Guid.Parse("{00024501-0000-0000-C000-000000000046}");
        var classIDs = new Dictionary<Guid, string>
        {
            [classID] = "Excel.Application",
        };
        var registrations = new[]
        {
            new RegisteredProgID(
                classID,
                "",
                ["Excel.Application"]),
            new RegisteredProgID(
                classID,
                "{00020813-0000-0000-C000-000000000046}",
                ["Excel.Application.16"]),
            new RegisteredProgID(
                otherClassID,
                "{00020813-0000-0000-C000-000000000046}",
                ["Excel.Workbook"]),
            new RegisteredProgID(
                classID,
                "{420B2830-E718-11CF-893D-00A0C9054228}",
                ["Scripting.Dictionary"]),
        };

        var progIDs = TypeLibImporterService.SelectProgIDsForTypeLib(
            "{00020813-0000-0000-C000-000000000046}",
            classIDs,
            registrations);

        Assert.Equal("Excel.Application", progIDs["Excel.Application"]);
        Assert.Equal("Excel.Application", progIDs["Excel.Application.16"]);
        Assert.False(progIDs.ContainsKey("Excel.Workbook"));
        Assert.False(progIDs.ContainsKey("Scripting.Dictionary"));
    }
}
