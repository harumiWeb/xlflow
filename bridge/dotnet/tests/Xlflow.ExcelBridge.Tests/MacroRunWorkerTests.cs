using Xlflow.ExcelBridge.Workers;

namespace Xlflow.ExcelBridge.Tests;

public sealed class MacroRunWorkerTests
{
    [Fact]
    public void CreateStartInfoUsesInternalWorkerMode()
    {
        var startInfo = MacroRunWorkerProcess.CreateStartInfo("request.json", "result.json");

        Assert.False(startInfo.UseShellExecute);
        Assert.Contains(BridgeStartup.InternalRunFlag, startInfo.ArgumentList);
        Assert.Contains("--run-worker", startInfo.ArgumentList);
        Assert.Equal("request.json", startInfo.Environment["XLFLOW_WORKER_REQUEST_PATH"]);
        Assert.Equal("result.json", startInfo.Environment["XLFLOW_WORKER_RESULT_PATH"]);
    }

    [Fact]
    public void ActivateTargetProjectActivatesWorkbookAndSelectsProject()
    {
        var workbook = new FakeWorkbook();
        var vbProject = new FakeVbProject();
        var vbe = new FakeVbe();

        MacroRunWorker.ActivateTargetProject(workbook, vbProject, vbe);

        Assert.True(workbook.Activated);
        Assert.Same(vbProject, vbe.ActiveVBProject);
    }

    [Fact]
    public void ActivateTargetProjectStillActivatesWorkbookWhenProjectSelectionFails()
    {
        var workbook = new FakeWorkbook();
        var vbProject = new FakeVbProject();
        var vbe = new FakeVbe { ThrowOnSetActiveProject = true };

        MacroRunWorker.ActivateTargetProject(workbook, vbProject, vbe);

        Assert.True(workbook.Activated);
        Assert.Null(vbe.ActiveVBProject);
    }

    public sealed class FakeWorkbook
    {
        public bool Activated { get; private set; }

        public void Activate()
        {
            Activated = true;
        }
    }

    public sealed class FakeVbProject;

    public sealed class FakeVbe
    {
        public bool ThrowOnSetActiveProject { get; init; }

        public object? ActiveVBProject
        {
            get => _activeVbProject;
            set
            {
                if (ThrowOnSetActiveProject)
                {
                    throw new InvalidOperationException("selection failed");
                }

                _activeVbProject = value;
            }
        }

        private object? _activeVbProject;
    }
}
