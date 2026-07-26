using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class BuildArtifactPublisherTests
{
    [Fact]
    public void CreateStageUsesOutputParentAndRejectsOfficeLock()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "release", "Book.xlsm");
        var stage = BuildArtifactPublisher.CreateStage(output);
        try
        {
            Assert.Equal(Path.GetDirectoryName(output), Path.GetDirectoryName(stage.Directory));
            Assert.Equal("Book.xlsm", Path.GetFileName(stage.TemporaryWorkbookPath));
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(output)!, "~$Book.xlsm"), "locked");

            var error = Assert.Throws<BuildArtifactException>(() => BuildArtifactPublisher.CreateStage(output));
            Assert.Equal("build_output_busy", error.Code);
        }
        finally
        {
            Directory.Delete(stage.Directory, recursive: true);
        }
    }

    [Fact]
    public void CreateStageReportsOutputDirectoryFailure()
    {
        using var workspace = new TemporaryWorkspace();
        var blockedParent = Path.Combine(workspace.Path, "not-a-directory");
        File.WriteAllText(blockedParent, "file");

        var error = Assert.Throws<BuildArtifactException>(() => BuildArtifactPublisher.CreateStage(Path.Combine(blockedParent, "Book.xlsm")));

        Assert.Equal("build_output_directory_failed", error.Code);
    }

    [Fact]
    public void ValidateTemporaryArtifactRejectsMissingAndEmptyArtifacts()
    {
        using var workspace = new TemporaryWorkspace();
        var missing = Path.Combine(workspace.Path, "missing.xlsm");
        var missingError = Assert.Throws<BuildArtifactException>(() => BuildArtifactPublisher.ValidateTemporaryArtifact(missing));
        Assert.Equal("build_temporary_artifact_missing", missingError.Code);

        File.WriteAllText(missing, "");
        var emptyError = Assert.Throws<BuildArtifactException>(() => BuildArtifactPublisher.ValidateTemporaryArtifact(missing));
        Assert.Equal("build_temporary_artifact_missing", emptyError.Code);
    }

    [Fact]
    public void PublishAtomicallyCreatesAndReplacesWithoutDeletingPreviousOutput()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Release.xlsm");
        var initial = Path.Combine(workspace.Path, "initial.xlsm");
        File.WriteAllText(initial, "first");

        var publisher = new BuildArtifactPublisher();
        var created = publisher.Publish(initial, output);
        Assert.False(created.ReplacedExisting);
        Assert.Equal("atomic_create", created.Publication);
        Assert.Equal("first", File.ReadAllText(output));

        var replacement = Path.Combine(workspace.Path, "replacement.xlsm");
        File.WriteAllText(replacement, "second");
        var replaced = publisher.Publish(replacement, output);
        Assert.True(replaced.ReplacedExisting);
        Assert.Equal("atomic_replace", replaced.Publication);
        Assert.Equal("second", File.ReadAllText(output));
    }

    [Fact]
    public void PublishRejectsLockedOutputBeforeReplacement()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Release.xlsm");
        var temporary = Path.Combine(workspace.Path, "staged.xlsm");
        File.WriteAllText(output, "previous");
        File.WriteAllText(temporary, "next");
        using var held = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);

        var error = Assert.Throws<BuildArtifactException>(() => new BuildArtifactPublisher().Publish(temporary, output));

        Assert.Equal("build_output_busy", error.Code);
        Assert.Equal("previous", File.ReadAllText(output));
    }

    [Fact]
    public void PublishReplaceFailureKeepsExistingOutputBytes()
    {
        using var workspace = new TemporaryWorkspace();
        var output = Path.Combine(workspace.Path, "Release.xlsm");
        var temporary = Path.Combine(workspace.Path, "staged.xlsm");
        File.WriteAllText(output, "previous artifact");
        File.WriteAllText(temporary, "new artifact");
        var publisher = new BuildArtifactPublisher(atomicReplace: (_, _) => new IOException("simulated replace failure"));

        var error = Assert.Throws<BuildArtifactException>(() => publisher.Publish(temporary, output));

        Assert.Equal("build_output_replace_failed", error.Code);
        Assert.Equal("previous artifact", File.ReadAllText(output));
        Assert.Equal("new artifact", File.ReadAllText(temporary));
    }

    [Fact]
    public void CleanupReportsResidueWhenOwnedStageCannotBeDeleted()
    {
        using var workspace = new TemporaryWorkspace();
        var stage = Path.Combine(workspace.Path, ".xlflow-build-test");
        Directory.CreateDirectory(stage);
        var publisher = new BuildArtifactPublisher(_ => new IOException("denied"));

        var cleanup = publisher.Cleanup(stage);

        Assert.False(cleanup.Succeeded);
        Assert.Equal(stage, cleanup.ResidualPath);
        Assert.Equal("denied", cleanup.Error);
        Directory.Delete(stage, recursive: true);
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        internal TemporaryWorkspace()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "xlflow-build-artifact-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
