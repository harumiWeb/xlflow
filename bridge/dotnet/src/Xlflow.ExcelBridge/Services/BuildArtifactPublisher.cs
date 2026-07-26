namespace Xlflow.ExcelBridge.Services;

// Owns only the staging directory it creates beside the requested output.  The
// output itself is never removed as part of a failed build.
internal sealed class BuildArtifactPublisher
{
    private readonly Func<string, Exception?> _deleteDirectory;
    private readonly Func<string, string, Exception?> _atomicReplace;

    internal BuildArtifactPublisher(Func<string, Exception?>? deleteDirectory = null, Func<string, string, Exception?>? atomicReplace = null)
    {
        _deleteDirectory = deleteDirectory ?? DeleteDirectory;
        _atomicReplace = atomicReplace ?? AtomicReplace;
    }

    internal static BuildArtifactStage CreateStage(string outputPath)
    {
        var outputParent = Path.GetDirectoryName(outputPath);
        if (string.IsNullOrWhiteSpace(outputParent))
        {
            throw new BuildArtifactException("build_output_directory_failed", "output workbook parent is required", "output_directory");
        }

        try
        {
            Directory.CreateDirectory(outputParent);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            throw new BuildArtifactException("build_output_directory_failed", "could not create the output workbook directory", "output_directory", ex);
        }

        EnsureOutputAvailable(outputPath);
        var stagingDirectory = Path.Combine(outputParent, ".xlflow-build-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(stagingDirectory);
            return new BuildArtifactStage(stagingDirectory, Path.Combine(stagingDirectory, Path.GetFileName(outputPath)));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            throw new BuildArtifactException("build_output_directory_failed", "could not create the build staging directory", "output_directory", ex);
        }
    }

    internal static void ValidateTemporaryArtifact(string temporaryWorkbook)
    {
        try
        {
            var info = new FileInfo(temporaryWorkbook);
            if (!info.Exists || info.Length == 0)
            {
                throw new BuildArtifactException("build_temporary_artifact_missing", "the staged workbook is missing or empty after Excel exited", "temporary_artifact");
            }

            using var stream = new FileStream(temporaryWorkbook, FileMode.Open, FileAccess.Read, FileShare.Read);
            _ = stream.Length;
        }
        catch (BuildArtifactException) { throw; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new BuildArtifactException("build_temporary_artifact_missing", "the staged workbook cannot be read after Excel exited", "temporary_artifact", ex);
        }
    }

    internal BuildArtifactPublication Publish(string temporaryWorkbook, string outputPath)
    {
        EnsureOutputAvailable(outputPath);
        try
        {
            if (!File.Exists(outputPath))
            {
                // The staging directory is under output's parent, so this is a
                // same-volume atomic create.  Do not overwrite a concurrent
                // creator: that is a publication failure, not a fallback case.
                File.Move(temporaryWorkbook, outputPath, overwrite: false);
                return new BuildArtifactPublication(false, "atomic_create");
            }

            // Replace is the Windows atomic replacement API.  In particular,
            // never degrade to delete-then-move: that would destroy the prior
            // successful artifact if publication failed halfway through.
            var replaceError = _atomicReplace(temporaryWorkbook, outputPath);
            if (replaceError is not null)
            {
                throw new BuildArtifactException("build_output_replace_failed", "could not atomically publish the staged workbook; the previous output was left unchanged", "publish", replaceError);
            }
            return new BuildArtifactPublication(true, "atomic_replace");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException or NotSupportedException)
        {
            throw new BuildArtifactException("build_output_replace_failed", "could not atomically publish the staged workbook; the previous output was left unchanged", "publish", ex);
        }
    }

    internal BuildArtifactCleanup Cleanup(string stagingDirectory)
    {
        var error = _deleteDirectory(stagingDirectory);
        return error is null
            ? new BuildArtifactCleanup(true, null, null)
            : new BuildArtifactCleanup(false, stagingDirectory, error.Message);
    }

    private static void EnsureOutputAvailable(string outputPath)
    {
        var lockPath = Path.Combine(Path.GetDirectoryName(outputPath)!, "~$" + Path.GetFileName(outputPath));
        if (File.Exists(lockPath))
        {
            throw new BuildArtifactException("build_output_busy", "the output workbook is open in Office", "output_busy");
        }

        if (!File.Exists(outputPath))
        {
            return;
        }

        try
        {
            using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new BuildArtifactException("build_output_busy", "the output workbook cannot be replaced because it is in use", "output_busy", ex);
        }
    }

    private static Exception? DeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
            return null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return ex;
        }
    }

    private static Exception? AtomicReplace(string sourcePath, string destinationPath)
    {
        try
        {
            File.Replace(sourcePath, destinationPath, destinationBackupFileName: null);
            return null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException or NotSupportedException)
        {
            return ex;
        }
    }
}

internal sealed record BuildArtifactStage(string Directory, string TemporaryWorkbookPath);

internal sealed record BuildArtifactPublication(bool ReplacedExisting, string Publication);

internal sealed record BuildArtifactCleanup(bool Succeeded, string? ResidualPath, string? Error);

internal sealed class BuildArtifactException(string code, string message, string stage, Exception? innerException = null) : Exception(message, innerException)
{
    internal string Code { get; } = code;
    internal string Stage { get; } = stage;
}
