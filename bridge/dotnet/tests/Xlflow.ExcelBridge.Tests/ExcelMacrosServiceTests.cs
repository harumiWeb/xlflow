using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class ExcelMacrosServiceTests
{
    [Fact]
    public void FindMacroProceduresDiscoversUnicodeNames()
    {
        const string code = """
            Option Explicit

            Public Sub Run()
            End Sub

            Sub 集計結果を作成()
            End Sub

            Sub Generate(path As String, count As Long)
            End Sub

            Public Function Build() As Boolean
            End Function

            private Static Sub Hidden()
            End Sub
            """;

        var macros = ExcelMacrosService.FindMacroProcedures("Main", "standard", code);

        Assert.Equal(5, macros.Count);
        Assert.Equal("Main.Run", macros[0].QualifiedName);
        Assert.True(macros[0].Runnable);
        Assert.Equal("Main.集計結果を作成", macros[1].QualifiedName);
        Assert.True(macros[1].Runnable);
        Assert.Equal("Generate", macros[2].Name);
        Assert.True(macros[2].HasParameters);
        Assert.False(macros[2].Runnable);
        Assert.Equal("has_parameters", macros[2].ReasonNotRunnable);
        Assert.Equal("Build", macros[3].Name);
        Assert.Equal("function", macros[3].Kind);
        Assert.Equal("Hidden", macros[4].Name);
        Assert.Equal("Private", macros[4].Visibility);
        Assert.True(macros[4].Runnable);
    }

    [Fact]
    public void FindMacroProceduresMarksNonEventDocumentProceduresRunnableRegardlessOfVisibility()
    {
        const string code = """
            Private Sub PrivateProbe()
            End Sub

            Public Static Sub PublicProbe()
            End Sub

            Friend Sub FriendProbe()
            End Sub

            Private Sub Worksheet_Activate()
            End Sub
            """;

        var macros = ExcelMacrosService.FindMacroProcedures("Sheet1", "document_module", code);

        Assert.Equal(4, macros.Count);
        Assert.Equal(["Private", "Public", "Friend", "Private"], macros.Select(m => m.Visibility).ToArray());
        Assert.All(macros.Take(3), macro => Assert.True(macro.Runnable));
        Assert.Equal("event_procedure", macros[3].ReasonNotRunnable);
        Assert.False(macros[3].Runnable);
    }
}
