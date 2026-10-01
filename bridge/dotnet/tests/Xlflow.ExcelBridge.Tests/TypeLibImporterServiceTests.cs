using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class TypeLibImporterServiceTests
{
    [Fact]
    public void ResolveLibrariesIncludesOutlookTypeLib()
    {
        var resolved = TypeLibImporterService.ResolveLibraries("outlook");
        var target = Assert.Single(resolved.Targets);

        Assert.Equal("Outlook", target.Name);
        Assert.Equal("{00062FFF-0000-0000-C000-000000000046}", target.LibID);
        Assert.Equal("outlook.generated.json", target.Output);
        Assert.False(resolved.BestEffort);
    }

    [Theory]
    [InlineData("1.9", 1, 9)]
    [InlineData("8.7", 8, 7)]
    [InlineData("2.c", 2, 12)]
    [InlineData("c.0", 12, 0)]
    public void ParseVersionUsesTypeLibHexadecimalSemantics(string value, int major, int minor)
    {
        var parsed = TypeLibImporterService.ParseVersion(value);

        Assert.NotNull(parsed);
        Assert.Equal(major, parsed.Value.Major);
        Assert.Equal(minor, parsed.Value.Minor);
    }

    [Theory]
    [InlineData("")]
    [InlineData("1")]
    [InlineData("1.2.3")]
    [InlineData("g.0")]
    [InlineData("10000.0")]
    public void ParseVersionRejectsInvalidRegistryVersions(string value)
    {
        Assert.Null(TypeLibImporterService.ParseVersion(value));
    }

    [Fact]
    public void ResolveLibrariesIncludesExpandedCatalogAndDaoFallback()
    {
        var resolved = TypeLibImporterService.ResolveLibraries("all");
        var byOutput = resolved.Targets.ToDictionary(target => target.Output, StringComparer.OrdinalIgnoreCase);

        Assert.True(resolved.BestEffort);
        Assert.Equal(16, resolved.Targets.Count);
        Assert.Equal(
            ["{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", "{00025E01-0000-0000-C000-000000000046}"],
            byOutput["dao.generated.json"].LibIDs);
        Assert.Equal("{F5078F18-C551-11D3-89B9-0000F81FE221}", byOutput["msxml.generated.json"].LibID);
        Assert.Equal("{662901FC-6951-4854-9EB2-D9A2570F2B2E}", byOutput["winhttp.generated.json"].LibID);
        Assert.Equal("{00020905-0000-0000-C000-000000000046}", byOutput["word.generated.json"].LibID);
        Assert.Equal("{91493440-5A91-11CF-8700-00AA0060263B}", byOutput["powerpoint.generated.json"].LibID);
        Assert.Equal("{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", byOutput["access.generated.json"].LibID);
        Assert.Equal("{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}", byOutput["wsh.generated.json"].LibID);
        Assert.Equal("{565783C6-CB41-11D1-8B02-00600806D9B6}", byOutput["wmi.generated.json"].LibID);
        Assert.Equal("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", byOutput["vbscript-regexp.generated.json"].LibID);
    }

    [Fact]
    public void CanonicalTypeNameUsesContainingLibraryQualifier()
    {
        Assert.Equal("Office.CommandBar", TypeLibImporterService.CanonicalTypeName("Office", "_CommandBar"));
        Assert.Equal("MSForms.CommandButton", TypeLibImporterService.CanonicalTypeName("MSForms", "CommandButton"));
    }

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

    [Fact]
    public void ResolveDefaultMemberSelectsOneNormalizedCandidate()
    {
        var resolved = TypeLibImporterService.ResolveDefaultMember(
        [
            new(" Item ", " Variant "),
            new("item", "variant"),
        ]);

        Assert.Equal(new("Item", "Variant"), resolved);
    }

    [Fact]
    public void ResolveDefaultMemberDoesNotChooseBetweenCandidateNames()
    {
        var resolved = TypeLibImporterService.ResolveDefaultMember(
        [
            new("Item", "Variant"),
            new("Value", "Variant"),
        ]);

        Assert.Null(resolved);
    }

    [Fact]
    public void ResolveDefaultMemberDoesNotChooseBetweenReturnTypes()
    {
        var resolved = TypeLibImporterService.ResolveDefaultMember(
        [
            new("Item", "Variant"),
            new("Item", "Object"),
        ]);

        Assert.Null(resolved);
    }

    [Fact]
    public void ResolveDefaultMemberIgnoresIncompleteCandidates()
    {
        var resolved = TypeLibImporterService.ResolveDefaultMember(
        [
            new("Item", ""),
            new(" Item ", "Variant"),
        ]);

        Assert.Equal(new("Item", "Variant"), resolved);
    }
}
