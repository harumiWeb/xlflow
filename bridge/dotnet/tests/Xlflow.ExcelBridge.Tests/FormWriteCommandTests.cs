using System.Collections;
using System.Reflection;
using System.Text;
using System.Text.Json;
using Xlflow.ExcelBridge.Commands;
using Xlflow.ExcelBridge.Contract;
using Xlflow.ExcelBridge.Serialization;
using Xlflow.ExcelBridge.Services;

namespace Xlflow.ExcelBridge.Tests;

public sealed class FormWriteCommandTests
{
    [Fact]
    public void HandleParsesPayloadAndReturnsExpectedExtensions()
    {
        var service = new FakeFormWriteService((request, args) =>
        {
            Assert.Equal("form-write", request.Command);
            Assert.Equal("build", args.Action);
            Assert.Equal(@"C:\work\book.xlsm", args.WorkbookPath);
            Assert.Equal(@"src\forms\specs\UserForm1.yaml", args.SpecPath);
            Assert.Equal(@"C:\work\src\forms", args.FormsDir);
            Assert.Equal("sidecar", args.CodeSource);
            Assert.True(args.Folders);
            Assert.Equal("update", args.FolderAnnotation);
            Assert.True(args.DefaultComponentFolders);
            Assert.Equal("eyJmb28iOiJiYXIifQ==", args.SpecJson64);
            Assert.True(args.Overwrite);
            Assert.False(args.NoSave);
            Assert.False(args.Visible);
            Assert.True(args.UseSession);
            Assert.Equal(@"C:\work\.xlflow\session.json", args.MetadataPath);

            return BridgeResponse.Ok(request, new Dictionary<string, object?>
            {
                ["target"] = new Dictionary<string, object?> { ["kind"] = "live_session", ["path"] = args.WorkbookPath },
                ["session"] = new Dictionary<string, object?> { ["active"] = true, ["workbook_path"] = args.WorkbookPath },
                ["workbook"] = new Dictionary<string, object?> { ["path"] = args.WorkbookPath, ["session"] = true },
                ["forms"] = new Dictionary<string, object?>
                {
                    ["name"] = "UserForm1",
                    ["action"] = args.Action,
                    ["source_synced"] = true,
                },
            });
        });
        var command = new FormWriteCommand(service);
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-write",
            Command = "form-write",
            Payload = JsonDocument.Parse("""
                {
                  "Action": "build",
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "SpecPath": "src\\forms\\specs\\UserForm1.yaml",
                  "FormsDir": "C:\\work\\src\\forms",
                  "CodeSource": "sidecar",
                  "Folders": "true",
                  "FolderAnnotation": "update",
                  "DefaultComponentFolders": "true",
                  "SpecJson64": "eyJmb28iOiJiYXIifQ==",
                  "Overwrite": "true",
                  "NoSave": "false",
                  "Visible": "false",
                  "UseSession": "true",
                  "MetadataPath": "C:\\work\\.xlflow\\session.json"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("ok", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("UserForm1", json.RootElement.GetProperty("forms").GetProperty("name").GetString());
        Assert.Equal("build", json.RootElement.GetProperty("forms").GetProperty("action").GetString());
    }

    [Fact]
    public void HandleRejectsInvalidNoSaveUsageBeforeService()
    {
        var command = new FormWriteCommand(new FakeFormWriteService((request, _) => BridgeResponse.Ok(request)));
        var request = new BridgeRequest
        {
            ProtocolVersion = ProtocolVersion.Current,
            RequestId = "req-form-write-invalid",
            Command = "form-write",
            Payload = JsonDocument.Parse("""
                {
                  "Action": "build",
                  "WorkbookPath": "C:\\work\\book.xlsm",
                  "FormsDir": "C:\\work\\src\\forms",
                  "SpecJson64": "eyJmb28iOiJiYXIifQ==",
                  "NoSave": "true",
                  "UseSession": "false"
                }
                """).RootElement.Clone(),
        };

        var response = command.Handle(request, CancellationToken.None);
        var json = JsonSerializer.SerializeToDocument(response, JsonOptions.Default);

        Assert.Equal("failed", json.RootElement.GetProperty("status").GetString());
        Assert.Equal("form_build_args_invalid", json.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void SetVBComponentPropertyUsesOneBasedVBIDEPropertyIndexes()
    {
        var component = new FakeVBComponent();
        var method = typeof(ExcelFormWriteService).GetMethod("SetVBComponentProperty", BindingFlags.NonPublic | BindingFlags.Static);

        Assert.NotNull(method);
        var result = method.Invoke(null, [component, "Width", 333.0]);

        Assert.Equal(true, result);
        Assert.Equal(333.0, component.Properties.Width.Value);
    }

    [Fact]
    public void SetRequiredVBComponentPropertyPersistsCaptionAndVerifiesReadback()
    {
        var component = new FakeVBComponent();
        var method = typeof(ExcelFormWriteService).GetMethod("SetRequiredVBComponentProperty", BindingFlags.NonPublic | BindingFlags.Static);

        Assert.NotNull(method);
        _ = method.Invoke(null, [component, "Caption", "労務管理 実行"]);

        Assert.Equal("労務管理 実行", component.Properties.Caption.Value);
    }

    [Fact]
    public void SetRequiredVBComponentPropertyFailsWhenCaptionReadbackDoesNotMatch()
    {
        var component = new FakeVBComponent(ignoreCaptionWrites: true);
        var method = typeof(ExcelFormWriteService).GetMethod("SetRequiredVBComponentProperty", BindingFlags.NonPublic | BindingFlags.Static);

        Assert.NotNull(method);
        var exception = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, [component, "Caption", "労務管理 実行"]));

        var inner = Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Contains("Caption did not persist", inner.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AddDesignerControlUsesPagesForMultiPageAndSetsValueAfterPageChildren()
    {
        const string json = """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Tabs","selectedIndex":1},{"id":"first","parentId":"multi","type":"Page","name":"PageA","caption":"First","tag":"page-tag","controlTipText":"page-tip","accelerator":"P","enabled":false,"visible":false},{"id":"second","parentId":"multi","type":"Page","name":"PageB"},{"id":"label","parentId":"first","type":"Label","name":"LabelA","left":12.5}]}""";
        var designer = new FakeFormDesigner();

        AddFirstControl(json, designer);

        var multiPage = Assert.IsType<FakeMultiPage>(Assert.Single(designer.Controls.Items));
        Assert.Equal(new[] { "PageA", "PageB" }, multiPage.Pages.AddedNames);
        Assert.Equal(new[] { "DefaultPage2", "DefaultPage1" }, multiPage.Pages.RemovedNames);
        Assert.Equal(new[] { "PageA", "PageB" }, multiPage.Pages.Items.Select(page => page.Name));
        Assert.Equal(1, multiPage.Value);
        Assert.Equal(2, multiPage.ValueSetPageCount);
        var page = multiPage.Pages.Items[0];
        Assert.Equal("First", page.Caption);
        Assert.Equal("page-tag", page.Tag);
        Assert.Equal("page-tip", page.ControlTipText);
        Assert.Equal("P", page.Accelerator);
        Assert.False(page.Enabled);
        Assert.False(page.Visible);
        Assert.Equal(12.5, Assert.Single(page.Controls.Items).Left);
        Assert.Empty(multiPage.Controls.Items);
    }

    [Fact]
    public void AddDesignerControlWritesPageMetadataFromCaseInsensitivePropertyBagAliases()
    {
        const string json = """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Tabs"},{"id":"page","parentId":"multi","type":"Page","name":"PageAlpha","caption":"Alpha","properties":{"CAPTION":"Alpha","ENABLED":false,"Visible":false,"TaG":"alpha-tag","ControlTipText":"alpha-tip","Accelerator":"A"}}]}""";
        var designer = new FakeFormDesigner();

        AddFirstControl(json, designer);

        var multiPage = Assert.IsType<FakeMultiPage>(Assert.Single(designer.Controls.Items));
        var page = Assert.Single(multiPage.Pages.Items);
        Assert.Equal("Alpha", page.Caption);
        Assert.False(page.Enabled);
        Assert.False(page.Visible);
        Assert.Equal("alpha-tag", page.Tag);
        Assert.Equal("alpha-tip", page.ControlTipText);
        Assert.Equal("A", page.Accelerator);
    }

    [Fact]
    public void DecodeSpecRejectsConflictingPageAliasesBeforeAnyControlMutation()
    {
        const string json = """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Tabs"},{"id":"page","parentId":"multi","type":"Page","name":"PageAlpha","tag":"top-level-tag","properties":{"TAG":"bag-tag"}}]}""";
        var designer = new FakeFormDesigner();

        var exception = Assert.Throws<TargetInvocationException>(() => AddFirstControl(json, designer));

        Assert.Contains("conflicting explicit aliases for 'tag'", exception.InnerException?.Message, StringComparison.Ordinal);
        Assert.Empty(designer.Controls.Items);
    }

    [Fact]
    public void DecodeSpecRejectsInvalidKnownSelectionRanges()
    {
        var invalidSpecs = new[]
        {
            """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Pages","selectedIndex":-1},{"id":"page","parentId":"multi","type":"Page","name":"PageA"}]}""",
            """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Pages","selectedIndex":1},{"id":"page","parentId":"multi","type":"Page","name":"PageA"}]}""",
            """{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","selectedIndex":-1,"tabs":[{"name":"TabA"}]}]}""",
            """{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","selectedIndex":1,"tabs":[{"name":"TabA"}]}]}""",
            """{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","selectedIndex":0,"tabs":[]}]}""",
        };

        foreach (var json in invalidSpecs)
        {
            var exception = Assert.Throws<TargetInvocationException>(() => DecodeSpecObject(json));
            Assert.Contains("selectedIndex must be", exception.InnerException?.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ApplySelectionPreflightUsesExistingTabCountWhenTabsAreOmitted()
    {
        var spec = DecodeSpecObject("""{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","selectedIndex":2}]}""");
        var designer = new FakeFormDesigner();
        var strip = new FakeTabStrip("Tabs", designer);
        strip.Tabs.Add("TabBeta");
        designer.Controls.Items.Add(strip);
        var setterCount = strip.ValueSetterCount;

        var exception = Assert.Throws<TargetInvocationException>(() => PrepareApplySelectionIndices(designer, spec));

        Assert.Contains("selectedIndex must be 0..1", exception.InnerException?.Message, StringComparison.Ordinal);
        Assert.Same(strip, Assert.Single(designer.Controls.Items));
        Assert.Equal(new[] { "DefaultTab", "TabBeta" }, strip.Tabs.Items.Select(tab => tab.Name));
        Assert.Equal(setterCount, strip.ValueSetterCount);
    }

    [Fact]
    public void ApplySelectionPreflightRetainsExistingTabsWhenAuthoredTabsAreOmitted()
    {
        var spec = DecodeSpecObject("""{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs"}]}""");
        var designer = new FakeFormDesigner();
        var strip = new FakeTabStrip("Tabs", designer);
        strip.Tabs.Add("TabBeta");
        strip.Tabs.Items[0].Caption = "Existing alpha";
        strip.Tabs.Items[1].Tag = "existing-beta";
        designer.Controls.Items.Add(strip);

        PrepareApplySelectionIndices(designer, spec);

        var firstControl = ((IEnumerable)GetDecodedControls(spec)).Cast<object>().Single();
        var observed = firstControl.GetType().GetProperty("Observed")?.GetValue(firstControl);
        Assert.NotNull(observed);
        var tabs = Assert.IsAssignableFrom<IEnumerable>(observed.GetType().GetProperty("Tabs")?.GetValue(observed));
        var tabNames = tabs.Cast<object>().Select(tab => tab.GetType().GetProperty("Name")?.GetValue(tab)).ToArray();
        Assert.Equal(new object?[] { "DefaultTab", "TabBeta" }, tabNames);
        var selectedIndex = observed.GetType().GetProperty("SelectedIndex")?.GetValue(observed);
        Assert.Equal(0, selectedIndex);
        Assert.Same(strip, Assert.Single(designer.Controls.Items));
    }

    [Fact]
    public void AddDesignerControlClearsDefaultPagesForAnAuthoredEmptyMultiPageWithoutSettingValue()
    {
        const string json = """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Tabs","selectedIndex":-1}]}""";
        var designer = new FakeFormDesigner();

        AddFirstControl(json, designer);

        var multiPage = Assert.IsType<FakeMultiPage>(Assert.Single(designer.Controls.Items));
        Assert.Empty(multiPage.Pages.Items);
        Assert.Equal(new[] { "DefaultPage2", "DefaultPage1" }, multiPage.Pages.RemovedNames);
        Assert.False(multiPage.EmptyValueSetterCalled);
    }

    [Fact]
    public void AddDesignerControlPreservesOmittedTabsAndClearsExplicitEmptyTabsByNumericIndex()
    {
        var omittedDesigner = new FakeFormDesigner();
        AddFirstControl("""{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs"}]}""", omittedDesigner);
        var omittedStrip = Assert.IsType<FakeTabStrip>(Assert.Single(omittedDesigner.Controls.Items));
        Assert.Equal(new[] { "DefaultTab" }, omittedStrip.Tabs.Items.Select(tab => tab.Name));
        Assert.Empty(omittedStrip.Tabs.RemoveArguments);

        var emptyDesigner = new FakeFormDesigner();
        AddFirstControl("""{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","tabs":[]}]}""", emptyDesigner);
        var emptyStrip = Assert.IsType<FakeTabStrip>(Assert.Single(emptyDesigner.Controls.Items));
        Assert.Empty(emptyStrip.Tabs.Items);
        Assert.Equal(new object[] { 0 }, emptyStrip.Tabs.RemoveArguments);
        Assert.Equal(-1, emptyStrip.Value);
        Assert.Equal(0, emptyStrip.ValueSetterCount);
    }

    [Fact]
    public void AddDesignerControlWritesTabMetadataAndSelectionThroughValueAfterTabsExist()
    {
        const string json = """{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","selectedIndex":1,"tabs":[{"name":"TabAlpha","caption":"Alpha","controlTipText":"tip","tag":"tag","accelerator":"A","enabled":false,"visible":true},{"name":"TabBeta"}]}]}""";
        var designer = new FakeFormDesigner();

        AddFirstControl(json, designer);

        var strip = Assert.IsType<FakeTabStrip>(Assert.Single(designer.Controls.Items));
        Assert.Equal(new object[] { 0 }, strip.Tabs.RemoveArguments);
        Assert.Equal(new[] { "TabAlpha", "TabBeta" }, strip.Tabs.Items.Select(tab => tab.Name));
        var first = strip.Tabs.Items[0];
        Assert.Equal("Alpha", first.Caption);
        Assert.Equal("tip", first.ControlTipText);
        Assert.Equal("tag", first.Tag);
        Assert.Equal("A", first.Accelerator);
        Assert.False(first.Enabled);
        Assert.True(first.Visible);
        Assert.Equal("default-caption", strip.Tabs.Items[1].Caption);
        Assert.Equal(1, strip.Value);
        Assert.Equal(2, strip.ValueSetTabCount);
    }

    [Fact]
    public void DecodeSpecRejectsTabWithoutRequiredName()
    {
        var decode = typeof(ExcelFormWriteService).GetMethod("DecodeSpec", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(decode);
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes("""{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"Tabs","tabs":[{}]}]}"""));

        var exception = Assert.Throws<TargetInvocationException>(() => decode!.Invoke(null, [encoded]));

        Assert.Contains("tabs[0].name is required", exception.InnerException?.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("{\"path\":\"src/forms/assets/logo.bmp\"}")]
    [InlineData("null")]
    public void DecodeSpecRejectsPictureAuthoringBeforeDesignerAccess(string pictureJson)
    {
        var decode = typeof(ExcelFormWriteService).GetMethod("DecodeSpec", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(decode);
        var json = "{\"form\":{\"name\":\"Sample\"},\"controls\":[{\"type\":\"Image\",\"name\":\"Logo\",\"picture\":" + pictureJson + "}]}";
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

        var exception = Assert.Throws<TargetInvocationException>(() => decode!.Invoke(null, [encoded]));

        Assert.Contains("controls[*].picture is supported only by pure-Go generation", exception.InnerException?.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DecodeSpecRejectsNestedPictureAuthoringBeforeDesignerAccess()
    {
        var decode = typeof(ExcelFormWriteService).GetMethod("DecodeSpec", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(decode);
        const string json = """{"form":{"name":"Sample"},"controls":[{"type":"Frame","name":"Frame","controls":[{"type":"Image","name":"Logo","picture":{"remove":true}}]}]}""";
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

        var exception = Assert.Throws<TargetInvocationException>(() => decode!.Invoke(null, [encoded]));

        Assert.Contains("controls[*].picture is supported only by pure-Go generation", exception.InnerException?.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RequiredModeledPropertyFailsWhenExcelDoesNotPersistIt()
    {
        var method = typeof(ExcelFormWriteService).GetMethod("SetRequiredMember", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);

        var exception = Assert.Throws<TargetInvocationException>(() => method!.Invoke(null, [new FakeUnpersistedProperty(), "Width", 24.0]));

        var inner = Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Contains("Width did not persist", inner.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PageObservedGeometryIsNotWrittenButChildGeometryIsApplied()
    {
        const string json = """{"form":{"name":"Sample"},"controls":[{"id":"multi","type":"MultiPage","name":"Tabs","width":300},{"id":"page","parentId":"multi","type":"Page","name":"PageAlpha","observed":{"left":4,"top":20,"width":210,"height":100}},{"parentId":"page","type":"TextBox","name":"InnerText","left":12.5,"width":80}]}""";
        var designer = new FakeFormDesigner();
        AddFirstControl(json, designer);
        var multi = Assert.IsType<FakeMultiPage>(Assert.Single(designer.Controls.Items));
        var page = Assert.Single(multi.Pages.Items);
        Assert.Equal(0, page.Left);
        Assert.Equal(0, page.Top);
        Assert.Equal(0, page.Width);
        Assert.Equal(0, page.Height);
        Assert.Equal(12.5, Assert.Single(page.Controls.Items).Left);
        Assert.Equal(80, Assert.Single(page.Controls.Items).Width);
    }

    [Fact]
    public void OmittedTabsAreCapturedThroughContainersWithoutProgIds()
    {
        var designer = new FakeFormDesigner();
        var multi = new FakeMultiPage("Pages", designer) { MissingProgId = true };
        var page = multi.Pages.Items[0];
        page.MissingProgId = true;
        var frame = new FakeFrame("Frame", page) { MissingProgId = true };
        var strip = new FakeTabStrip("NestedTabs", frame) { MissingProgId = true };
        strip.Tabs.Add("SecondTab");
        strip.Value = 1;
        designer.Controls.Items.Add(multi);
        page.Controls.Items.Add(frame);
        frame.Controls.Items.Add(strip);
        foreach (var (control, type) in new (FakeControl, string)[] { (multi, "MultiPage"), (page, "Page"), (frame, "Frame"), (strip, "TabStrip") })
        {
            System.ComponentModel.TypeDescriptor.AddProvider(new ControlTypeProvider(type), control);
        }
        var spec = DecodeSpecObject("""{"form":{"name":"Sample"},"controls":[{"type":"TabStrip","name":"NestedTabs"}]}""");
        PrepareApplySelectionIndices(designer, spec);
        var controlSpec = ((IEnumerable)GetDecodedControls(spec)).Cast<object>().Single();
        var observed = controlSpec.GetType().GetProperty("Observed")!.GetValue(controlSpec)!;
        var tabs = (IEnumerable)observed.GetType().GetProperty("Tabs")!.GetValue(observed)!;
        Assert.Equal(new[] { "DefaultTab", "SecondTab" }, tabs.Cast<object>().Select(tab => tab.GetType().GetProperty("Name")!.GetValue(tab)));
        Assert.Equal(1, observed.GetType().GetProperty("SelectedIndex")!.GetValue(observed));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TabOrderIsAppliedAfterAllSiblingsExist(bool nested)
    {
        var children = """[{"id":"first","type":"TextBox","name":"First","tabIndex":1,"zIndex":0},{"id":"second","type":"TextBox","name":"Second","tabIndex":0,"zIndex":1}]""";
        var json = nested
            ? "{\"form\":{\"name\":\"Sample\"},\"controls\":[{\"id\":\"frame\",\"type\":\"Frame\",\"name\":\"Frame\",\"controls\":" + children + "}]}"
            : "{\"form\":{\"name\":\"Sample\"},\"controls\":" + children + "}";
        var spec = DecodeSpecObject(json);
        var controls = GetDecodedControls(spec);
        var roots = typeof(ExcelFormWriteService).GetMethod("GetRootControls", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [spec]);
        var designer = new FakeFormDesigner();
        typeof(ExcelFormWriteService).GetMethod("AddDesignerControls", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [designer, roots, controls, "UserForm"]);
        var siblings = nested ? Assert.IsType<FakeFrame>(Assert.Single(designer.Controls.Items)).Controls.Items : designer.Controls.Items;
        Assert.Equal(new[] { "First", "Second" }, siblings.Select(control => control.Name));
        Assert.Equal(new[] { 1, 0 }, siblings.Select(control => control.TabIndex));
    }

    [Fact]
    public void ListItemsExistBeforeRequiredTextAndValueWrites()
    {
        var designer = new FakeFormDesigner();
        AddFirstControl("""{"form":{"name":"Sample"},"controls":[{"type":"ListBox","name":"List","list":["Alpha","Beta"],"value":"Beta","text":"Beta"}]}""", designer);
        var list = Assert.IsType<FakeListBox>(Assert.Single(designer.Controls.Items));
        Assert.Equal("Beta", list.Text);
        Assert.Equal("Beta", list.Value);
        Assert.Equal(new[] { "Alpha", "Beta" }, list.Items);
    }

    [Theory]
    [InlineData("quote \" text", "   Caption = \"quote \"\" text\"")]
    [InlineData("line\r\nEnd\r\nSub Injected()", null)]
    [InlineData("line\nEnd", null)]
    public void CaptionNormalizationCannotAddSourceLines(string caption, string? expected)
    {
        const string content = "VERSION 5.00\r\nBegin\r\n   Caption = \"Original\"\r\nEnd\r\nAttribute VB_Name = \"Sample\"";
        var result = (string)typeof(ExcelFormWriteService).GetMethod("InjectOrUpdateCaption", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [content, caption])!;
        if (expected is null)
        {
            Assert.Equal(content, result);
        }
        else
        {
            Assert.Contains(expected, result, StringComparison.Ordinal);
            Assert.Equal(content.Split('\n').Length, result.Split('\n').Length);
        }
    }

    private sealed class ControlTypeProvider(string type) : System.ComponentModel.TypeDescriptionProvider
    {
        public override System.ComponentModel.ICustomTypeDescriptor GetTypeDescriptor(Type objectType, object? instance) => new ControlTypeDescriptor(type);
    }

    private sealed class ControlTypeDescriptor(string type) : System.ComponentModel.CustomTypeDescriptor
    {
        public override string GetClassName() => type;
    }

    private static void AddFirstControl(string json, FakeFormDesigner designer)
    {
        var spec = DecodeSpecObject(json);
        var controls = GetDecodedControls(spec);
        var firstControl = ((IEnumerable)controls).Cast<object>().First();
        var add = typeof(ExcelFormWriteService).GetMethod("AddDesignerControl", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("AddDesignerControl method was not found.");
        add.Invoke(null, [designer, firstControl, controls, "UserForm"]);
    }

    private static object DecodeSpecObject(string json)
    {
        var decode = typeof(ExcelFormWriteService).GetMethod("DecodeSpec", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("DecodeSpec method was not found.");
        var encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        return decode.Invoke(null, [encoded]) ?? throw new InvalidOperationException("form spec was not decoded.");
    }

    private static object GetDecodedControls(object spec)
    {
        return spec.GetType().GetProperty("Controls")?.GetValue(spec)
            ?? throw new InvalidOperationException("form controls were not decoded.");
    }

    private static void PrepareApplySelectionIndices(FakeFormDesigner designer, object spec)
    {
        var method = typeof(ExcelFormWriteService).GetMethod("PrepareApplySelectionIndices", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("PrepareApplySelectionIndices method was not found.");
        method.Invoke(null, [designer, spec]);
    }

    private sealed class FakeFormWriteService(Func<BridgeRequest, FormWriteCommandArguments, BridgeResponse> handler) : IFormWriteService
    {
        public BridgeResponse Execute(BridgeRequest request, FormWriteCommandArguments args, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return handler(request, args);
        }
    }

    private sealed class FakeVBComponent
    {
        public FakeVBComponent(bool ignoreCaptionWrites = false)
        {
            Properties = new FakeVBIDEProperties(ignoreCaptionWrites);
        }

        public FakeVBIDEProperties Properties { get; }
    }

    private sealed class FakeVBIDEProperties
    {
        public FakeVBIDEProperties(bool ignoreCaptionWrites)
        {
            Caption = new FakeVBIDEProperty("Caption", "UserForm1", ignoreCaptionWrites);
        }

        public FakeVBIDEProperty Width { get; } = new("Width", 240.0);

        public FakeVBIDEProperty Height { get; } = new("Height", 180.0);

        public FakeVBIDEProperty Caption { get; }

        public int Count => 3;

        public FakeVBIDEProperty Item(int index) => index switch
        {
            1 => Width,
            2 => Height,
            3 => Caption,
            _ => throw new ArgumentOutOfRangeException(nameof(index), "VBIDE Properties is one-based."),
        };
    }

    private sealed class FakeVBIDEProperty(string name, object value, bool ignoreWrites = false)
    {
        private object _value = value;

        public string Name { get; } = name;

        public object Value
        {
            get => _value;
            set
            {
                if (!ignoreWrites)
                {
                    _value = value;
                }
            }
        }
    }

    private sealed class FakeFormDesigner
    {
        public FakeFormDesigner()
        {
            Controls = new FakeControlCollection(this);
        }

        public FakeControlCollection Controls { get; }
    }

    private class FakeControl(string name, string progId, object? parent)
    {
        public string Name { get; set; } = name;

        public bool MissingProgId { get; set; }

        public string ProgId => MissingProgId ? throw new InvalidOperationException("ProgId unavailable") : progId;

        public object? Parent { get; } = parent;

        public string Caption { get; set; } = "";

        public string Tag { get; set; } = "";

        public string ControlTipText { get; set; } = "";

        public string Accelerator { get; set; } = "";

        public virtual string Text { get; set; } = "";

        public double Left { get; set; }

        public double Top { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        private int _tabIndex = parent switch { FakeFormDesigner designer => designer.Controls.Count, FakeFrame frame => frame.Controls.Count, _ => 0 };

        public int TabIndex
        {
            get => _tabIndex;
            set
            {
                var siblings = Parent switch { FakeFormDesigner designer => designer.Controls.Items, FakeFrame frame => frame.Controls.Items, _ => null };
                if (siblings is null) { _tabIndex = value; return; }
                var next = Math.Clamp(value, 0, siblings.Count - 1);
                foreach (var sibling in siblings.Where(sibling => sibling != this))
                {
                    if (sibling._tabIndex >= next && sibling._tabIndex < _tabIndex) { sibling._tabIndex++; }
                    else if (sibling._tabIndex <= next && sibling._tabIndex > _tabIndex) { sibling._tabIndex--; }
                }
                _tabIndex = next;
            }
        }

        public bool Enabled { get; set; } = true;

        public bool Visible { get; set; } = true;

    }

    private sealed class FakeControlCollection(object owner)
    {
        public List<FakeControl> Items { get; } = [];

        public int Count => Items.Count;

        public object Item(int index) => Items[index];

        public object Item(string name) => Items.Single(control => control.Name == name);

        public object Add(string progId, string name, bool visible)
        {
            FakeControl control = progId switch
            {
                "Forms.MultiPage.1" => new FakeMultiPage(name, owner),
                "Forms.TabStrip.1" => new FakeTabStrip(name, owner),
                "Forms.Frame.1" => new FakeFrame(name, owner),
                "Forms.ListBox.1" => new FakeListBox(name, owner),
                _ => new FakeControl(name, progId, owner),
            };
            control.Visible = visible;
            Items.Add(control);
            return control;
        }
    }

    private sealed class FakeFrame : FakeControl
    {
        public FakeFrame(string name, object? parent) : base(name, "Forms.Frame.1", parent) { Controls = new FakeControlCollection(this); }
        public FakeControlCollection Controls { get; }
    }

    private sealed class FakeListBox(string name, object? parent) : FakeControl(name, "Forms.ListBox.1", parent)
    {
        public List<string> Items { get; } = [];
        public int ListCount => Items.Count;
        public void AddItem(string item) => Items.Add(item);
        public void Clear() => Items.Clear();
        public override string Text
        {
            get => base.Text;
            set { if (!Items.Contains(value)) { throw new InvalidOperationException("Text must match an existing item"); } base.Text = value; }
        }
        private string _value = "";
        public string Value
        {
            get => _value;
            set { if (!Items.Contains(value)) { throw new InvalidOperationException("Value must match an existing item"); } _value = value; }
        }
    }

    private sealed class FakeMultiPage(string name, object? parent) : FakeControl(name, "Forms.MultiPage.1", parent)
    {
        private int _value = 1;

        public FakeMultiPage() : this("MultiPage", null)
        {
        }

        public FakePageCollection Pages { get; } = new();

        public FakeControlCollection Controls { get; } = new(new object());

        public bool EmptyValueSetterCalled { get; private set; }

        public int ValueSetPageCount { get; private set; }

        public int Value
        {
            get => Pages.Count == 0 ? -1 : _value;
            set
            {
                if (Pages.Count == 0)
                {
                    EmptyValueSetterCalled = true;
                    throw new InvalidOperationException("Setting Value on an empty MultiPage is unsafe.");
                }
                if (value < 0 || value >= Pages.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(value));
                }
                _value = value;
                ValueSetPageCount = Pages.Count;
            }
        }
    }

    private sealed class FakePageCollection
    {
        public List<FakePage> Items { get; } = [new("DefaultPage1", null), new("DefaultPage2", null)];

        public List<string> AddedNames { get; } = [];

        public List<string> RemovedNames { get; } = [];

        public int Count => Items.Count;

        public object Item(int index) => Items[index];

        public object Add(string name)
        {
            AddedNames.Add(name);
            var page = new FakePage(name, null);
            Items.Add(page);
            return page;
        }

        public void Remove(string name)
        {
            RemovedNames.Add(name);
            var page = Items.Single(item => string.Equals(item.Name, name, StringComparison.Ordinal));
            Items.Remove(page);
        }
    }

    private sealed class FakePage(string name, object? parent) : FakeControl(name, "Forms.Page.1", parent)
    {
        public FakePage() : this("Page", null)
        {
        }

        public FakeControlCollection Controls { get; } = new(new object());
    }

    private sealed class FakeTabStrip(string name, object? parent) : FakeControl(name, "Forms.TabStrip.1", parent)
    {
        private int _value = 0;

        public FakeTabStrip() : this("TabStrip", null)
        {
        }

        public FakeTabs Tabs { get; } = new();

        public int ValueSetterCount { get; private set; }

        public int ValueSetTabCount { get; private set; }

        public int Value
        {
            get => Tabs.Count == 0 ? -1 : _value;
            set
            {
                if (value < 0 || value >= Tabs.Count)
                {
                    if (value == -1 && Tabs.Count == 0)
                    {
                        _value = value;
                        ValueSetterCount++;
                        return;
                    }
                    throw new ArgumentOutOfRangeException(nameof(value));
                }
                _value = value;
                ValueSetterCount++;
                ValueSetTabCount = Tabs.Count;
            }
        }
    }

    private sealed class FakeTabs
    {
        public List<FakeTab> Items { get; } = [new("DefaultTab")];

        public List<object> RemoveArguments { get; } = [];

        public int Count => Items.Count;

        public object Item(int index) => Items[index];

        public object Add(string name)
        {
            var tab = new FakeTab(name);
            Items.Add(tab);
            return tab;
        }

        public void Remove(int index)
        {
            RemoveArguments.Add(index);
            Items.RemoveAt(index);
        }
    }

    private sealed class FakeTab(string name)
    {
        public string Name { get; set; } = name;

        public string Caption { get; set; } = "default-caption";

        public string ControlTipText { get; set; } = "default-tip";

        public string Tag { get; set; } = "default-tag";

        public string Accelerator { get; set; } = "default-accelerator";

        public bool Enabled { get; set; } = true;

        public bool Visible { get; set; } = false;
    }

    private sealed class FakeUnpersistedProperty
    {
        public double Width
        {
            get => 0;
            set { }
        }
    }
}
