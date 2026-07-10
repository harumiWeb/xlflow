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
    public void ActivateTargetProjectUsesWorkbookActivationWhenAvailable()
    {
        var vbe = new FakeVbe();
        var vbProject = new FakeVbProject(vbe);
        var workbook = new FakeWorkbook(() => vbe.SetActiveProject(vbProject));

        MacroRunWorker.ActivateTargetProject(workbook, vbProject, vbe);

        Assert.True(workbook.Activated);
        Assert.Same(vbProject, vbe.ActiveVBProject);
        Assert.Equal(0, vbe.ActiveVBProjectSetCalls);
    }

    [Fact]
    public void ActivateTargetProjectShowsTargetCodePaneWhenWorkbookActivationFails()
    {
        var vbe = new FakeVbe();
        var codePane = new FakeCodePane(() => vbe.SetActiveProjectFromCodePane());
        var component = new FakeVbComponent(new FakeCodeModule(codePane), throwOnActivate: true);
        var vbProject = new FakeVbProject(vbe, component);
        vbe.CodePaneProject = vbProject;
        var workbook = new FakeWorkbook(() => { }, throwOnActivate: true);

        MacroRunWorker.ActivateTargetProject(workbook, vbProject, vbe);

        Assert.True(workbook.ActivateAttempted);
        Assert.True(component.ActivateAttempted);
        Assert.True(codePane.Shown);
        Assert.Same(vbProject, vbe.ActiveVBProject);
        Assert.Equal(0, vbe.ActiveVBProjectSetCalls);
    }

    [Fact]
    public void ActivateTargetProjectFailsWhenFallbackDoesNotSelectTargetProject()
    {
        var vbe = new FakeVbe();
        var codePane = new FakeCodePane(() => { });
        var component = new FakeVbComponent(new FakeCodeModule(codePane), throwOnActivate: true);
        var vbProject = new FakeVbProject(vbe, component);
        var workbook = new FakeWorkbook(() => { }, throwOnActivate: true);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            MacroRunWorker.ActivateTargetProject(workbook, vbProject, vbe));

        Assert.Contains("could not activate the target VBProject", ex.Message);
        Assert.True(workbook.ActivateAttempted);
        Assert.True(component.ActivateAttempted);
        Assert.True(codePane.Shown);
        Assert.Null(vbe.ActiveVBProject);
        Assert.Equal(0, vbe.ActiveVBProjectSetCalls);
    }

    public sealed class FakeWorkbook
    {
        private readonly Action _onActivate;
        private readonly bool _throwOnActivate;

        public FakeWorkbook(Action? onActivate = null, bool throwOnActivate = false)
        {
            _onActivate = onActivate ?? (() => { });
            _throwOnActivate = throwOnActivate;
        }

        public bool ActivateAttempted { get; private set; }

        public bool Activated { get; private set; }

        public void Activate()
        {
            ActivateAttempted = true;
            if (_throwOnActivate)
            {
                throw new InvalidOperationException("activation failed");
            }

            Activated = true;
            _onActivate();
        }
    }

    public sealed class FakeVbProject
    {
        public FakeVbProject(FakeVbe vbe, params FakeVbComponent[] components)
        {
            VBE = vbe;
            VBComponents = new FakeVbComponents(components);
        }

        public string Name => "VBAProject";

        public string FileName => "C:\\project\\selfmacros.xlam";

        public FakeVbe VBE { get; }

        public FakeVbComponents VBComponents { get; }
    }

    public sealed class FakeVbComponents(IReadOnlyList<FakeVbComponent> components)
    {
        public int Count => components.Count;

        public FakeVbComponent Item(int index)
        {
            return components[index - 1];
        }
    }

    public sealed class FakeVbComponent(FakeCodeModule codeModule, bool throwOnActivate = false)
    {
        public bool ActivateAttempted { get; private set; }

        public FakeCodeModule CodeModule { get; } = codeModule;

        public void Activate()
        {
            ActivateAttempted = true;
            if (throwOnActivate)
            {
                throw new InvalidOperationException("component activation failed");
            }
        }
    }

    public sealed class FakeCodeModule(FakeCodePane codePane)
    {
        public FakeCodePane CodePane { get; } = codePane;
    }

    public sealed class FakeCodePane(Action onShow)
    {
        public bool Shown { get; private set; }

        public void Show()
        {
            Shown = true;
            onShow();
        }
    }

    public sealed class FakeVbe
    {
        public int ActiveVBProjectSetCalls { get; private set; }

        public FakeVbProject? CodePaneProject { get; set; }

        public object? ActiveVBProject
        {
            get => _activeVbProject;
            set
            {
                ActiveVBProjectSetCalls++;
                throw new InvalidOperationException("ActiveVBProject is read-only");
            }
        }

        private object? _activeVbProject;

        public void SetActiveProject(object vbProject)
        {
            _activeVbProject = vbProject;
        }

        public void SetActiveProjectFromCodePane()
        {
            _activeVbProject = CodePaneProject;
        }
    }
}
