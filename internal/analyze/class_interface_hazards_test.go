package analyze

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

// runClassHazardProjectAnalysis builds an in-memory project and runs batch
// analysis with the requested hazard switches enabled.
func runClassHazardProjectAnalysis(t *testing.T, files []sourceproject.SourceFile, configure func(*config.AnalyzeConfig)) []Finding {
	t.Helper()
	cfg := config.Default()
	if configure != nil {
		configure(&cfg.Analyze)
	}
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: files})
	if err != nil {
		t.Fatal(err)
	}
	return result.Findings
}

func classFile(name, body string) sourceproject.SourceFile {
	return sourceproject.SourceFile{Path: name, Source: []byte(body), ModuleKind: sourceproject.ModuleKindClass}
}

func standardFile(name, body string) sourceproject.SourceFile {
	return sourceproject.SourceFile{Path: name, Source: []byte(body), ModuleKind: sourceproject.ModuleKindStandard}
}

func documentFile(name, body string) sourceproject.SourceFile {
	return sourceproject.SourceFile{Path: name, Source: []byte(body), ModuleKind: sourceproject.ModuleKindDocument}
}

func formFile(name, body string) sourceproject.SourceFile {
	return sourceproject.SourceFile{Path: name, Source: []byte(body), ModuleKind: sourceproject.ModuleKindForm}
}

// --- VBA278: public member underscore names -----------------------------

func TestVBA278FlagsPublicUnderscoredMemberInClass(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Sub Do_Work()\nEnd Sub\nPublic Sub Run()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	got := findingsByCode(findings, "VBA278")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Do_Work") {
		t.Fatalf("VBA278 findings = %+v, want one finding for Do_Work", got)
	}
	if got[0].Line != 2 {
		t.Fatalf("VBA278 line = %d, want the Sub declaration line 2", got[0].Line)
	}
}

func TestVBA278SkipsPrivateAndFriendMembers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPrivate Sub Helper_One()\nEnd Sub\nFriend Sub Helper_Two()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	if got := findingsByCode(findings, "VBA278"); len(got) != 0 {
		t.Fatalf("VBA278 findings = %+v, want none for private/friend members", got)
	}
}

func TestVBA278SkipsStandardModules(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		standardFile("Main.bas", "Option Explicit\nPublic Sub Do_Work()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	if got := findingsByCode(findings, "VBA278"); len(got) != 0 {
		t.Fatalf("VBA278 findings = %+v, want none in standard modules", got)
	}
}

func TestVBA278SkipsEventHandlers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\nPublic Sub Worksheet_Follow()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	for _, finding := range findingsByCode(findings, "VBA278") {
		if strings.Contains(finding.Message, "Worksheet_Change") {
			t.Fatalf("VBA278 flagged the event handler: %+v", finding)
		}
	}
}

func TestVBA278SkipsImplementedInterfaceMembers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("IFace.cls", "Option Explicit\nPublic Sub Go()\nEnd Sub\n"),
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPrivate Sub IFace_Go()\nEnd Sub\nPublic Sub IFace_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	for _, finding := range findingsByCode(findings, "VBA278") {
		if strings.Contains(finding.Message, "IFace_Go") {
			t.Fatalf("VBA278 flagged an implemented-interface member: %+v", finding)
		}
	}
}

func TestVBA278FlagsUnverifiedInterfacePrefixMember(t *testing.T) {
	// The resolved IFace declares no Utility member, so IFace_Utility is an
	// ordinary underscored helper rather than an interface binding.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("IFace.cls", "Option Explicit\nPublic Sub Go()\nEnd Sub\n"),
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPublic Sub IFace_Utility()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	got := findingsByCode(findings, "VBA278")
	if len(got) != 1 || !strings.Contains(got[0].Message, "IFace_Utility") {
		t.Fatalf("VBA278 findings = %+v, want one finding for the unverified IFace_Utility prefix", got)
	}
}

func TestVBA278SkipsUnresolvedInterfacePrefixMember(t *testing.T) {
	// IFoo is implemented but not part of the analyzed file set, so the
	// binding cannot be verified or disproved and both rules fail open.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nImplements IFoo\nPublic Sub IFoo_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	if got := findingsByCode(findings, "VBA278"); len(got) != 0 {
		t.Fatalf("VBA278 findings = %+v, want none for an unresolved Implements target", got)
	}
}

func TestVBA278FlagsPublicFormHelperWithUnknownControls(t *testing.T) {
	// Without a designer artifact the control set is unknown; only intrinsic
	// UserForm_* events are recognized, so Refresh_Data is a public helper.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		formFile("Dialog.frm", "Attribute VB_Name = \"Dialog\"\nOption Explicit\nPublic Sub Refresh_Data()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	got := findingsByCode(findings, "VBA278")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Refresh_Data") {
		t.Fatalf("VBA278 findings = %+v, want one finding for the form helper Refresh_Data", got)
	}
}

func TestVBA278SkipsIntrinsicUserFormEvent(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		formFile("Dialog.frm", "Attribute VB_Name = \"Dialog\"\nOption Explicit\nPublic Sub UserForm_Initialize()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	if got := findingsByCode(findings, "VBA278"); len(got) != 0 {
		t.Fatalf("VBA278 findings = %+v, want none for the intrinsic UserForm_Initialize event", got)
	}
}

func TestVBA278SkipsClassLifecycleMembers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Sub Class_Initialize()\nEnd Sub\nPublic Sub Class_Terminate()\nEnd Sub\nPublic Sub Other_Name()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	got := findingsByCode(findings, "VBA278")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Other_Name") {
		t.Fatalf("VBA278 findings = %+v, want a single Other_Name finding", got)
	}
}

// --- VBA279: public Enum in document modules -----------------------------

func TestVBA279FlagsPublicEnumInDocumentModule(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPublic Enum SheetMode\n  SheetRead\n  SheetWrite\nEnd Enum\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectDocumentModulePublicEnum = true })
	got := findingsByCode(findings, "VBA279")
	if len(got) != 1 || !strings.Contains(got[0].Message, "SheetMode") {
		t.Fatalf("VBA279 findings = %+v, want one finding for SheetMode", got)
	}
	if got[0].Line != 2 {
		t.Fatalf("VBA279 line = %d, want the Enum declaration line 2", got[0].Line)
	}
}

func TestVBA279SkipsPrivateEnumAndOtherModuleKinds(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate Enum SheetMode\n  SheetRead\nEnd Enum\n"),
		standardFile("Main.bas", "Option Explicit\nPublic Enum Mode\n  ModeRead\nEnd Enum\n"),
		classFile("Widget.cls", "Option Explicit\nPublic Enum WidgetMode\n  WidgetRead\nEnd Enum\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectDocumentModulePublicEnum = true })
	if got := findingsByCode(findings, "VBA279"); len(got) != 0 {
		t.Fatalf("VBA279 findings = %+v, want none for private/other-module enums", got)
	}
}

// --- VBA280: write-only property ----------------------------------------

func TestVBA280FlagsWriteOnlyProperty(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Size(v As Long)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	got := findingsByCode(findings, "VBA280")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Size") || !strings.Contains(got[0].Message, "write-only") {
		t.Fatalf("VBA280 findings = %+v, want one write-only finding for Size", got)
	}
}

func TestVBA280SatisfiedByMatchingGet(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Size(v As Long)\nEnd Property\nPublic Property Get Size() As Long\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA280"); len(got) != 0 {
		t.Fatalf("VBA280 findings = %+v, want none when a Get exists", got)
	}
}

func TestVBA280SkipsPrivateWriters(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPrivate Property Let Size(v As Long)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA280"); len(got) != 0 {
		t.Fatalf("VBA280 findings = %+v, want none for private writers", got)
	}
}

func TestVBA280FlagsFriendWriters(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nFriend Property Let Size(v As Long)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	got := findingsByCode(findings, "VBA280")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Size") {
		t.Fatalf("VBA280 findings = %+v, want one finding for the Friend writer", got)
	}
}

func TestVBA280FailsOpenOnConditionalGetter(t *testing.T) {
	// The only Property Get sits under a conditional-compilation branch, so
	// the IR cannot prove a matching getter in every configuration and the
	// property stays silent.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Size(v As Long)\nEnd Property\n#If DEBUG Then\nPublic Property Get Size() As Long\nEnd Property\n#End If\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA280"); len(got) != 0 {
		t.Fatalf("VBA280 findings = %+v, want none when the only Get is conditional", got)
	}
}

func TestVBA280DeduplicatesLetAndSet(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Value(v As Long)\nEnd Property\nPublic Property Set Value(v As Object)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA280"); len(got) != 1 {
		t.Fatalf("VBA280 findings = %+v, want one deduplicated finding", got)
	}
}

// --- VBA281: interface/event exposure ------------------------------------

func TestVBA281FlagsPublicImplementedInterfaceMember(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("IFace.cls", "Option Explicit\nPublic Sub Go()\nEnd Sub\n"),
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPublic Sub IFace_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA281")
	if len(got) != 1 || !strings.Contains(got[0].Message, "IFace_Go") {
		t.Fatalf("VBA281 findings = %+v, want one finding for IFace_Go", got)
	}
}

func TestVBA281SkipsUnverifiedInterfacePrefixMember(t *testing.T) {
	// IFace declares no Utility member, so IFace_Utility is not a verified
	// binding and produces no collision finding; VBA278 owns the plain
	// underscore-name diagnosis instead.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("IFace.cls", "Option Explicit\nPublic Sub Go()\nEnd Sub\n"),
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPublic Sub IFace_Utility()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA281"); len(got) != 0 {
		t.Fatalf("VBA281 findings = %+v, want none for an unverified interface prefix", got)
	}
}

func TestVBA281FailsOpenOnUnresolvedInterface(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nImplements IFoo\nPublic Sub IFoo_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA281"); len(got) != 0 {
		t.Fatalf("VBA281 findings = %+v, want none for an unresolved Implements target", got)
	}
}

func TestVBA281MatchesFullUnderscoredInterfaceName(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("I_Foo.cls", "Option Explicit\nPublic Sub Bar()\nEnd Sub\n"),
		classFile("Widget.cls", "Option Explicit\nImplements I_Foo\nPublic Sub I_Foo_Bar()\nEnd Sub\nPublic Sub I_Other()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA281")
	if len(got) != 1 || !strings.Contains(got[0].Message, "I_Foo_Bar") {
		t.Fatalf("VBA281 findings = %+v, want only I_Foo_Bar", got)
	}
}

func TestVBA281FlagsExplicitPublicEventHandler(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPublic Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA281")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Worksheet_Change") {
		t.Fatalf("VBA281 findings = %+v, want one finding for the explicit Public event handler", got)
	}
}

func TestVBA281SkipsImplicitVisibilityEventHandler(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\nSub Worksheet_SelectionChange(ByVal Target As Range)\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA281"); len(got) != 0 {
		t.Fatalf("VBA281 findings = %+v, want none without explicit Public", got)
	}
}

func TestVBA281SkipsPrivateImplementations(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("IFace.cls", "Option Explicit\nPublic Sub Go()\nEnd Sub\n"),
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPrivate Sub IFace_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA281"); len(got) != 0 {
		t.Fatalf("VBA281 findings = %+v, want none for private implementations", got)
	}
}

func TestVBA281SkipsFormHelperWhenControlsKnown(t *testing.T) {
	// With a supplied designer artifact listing only TextBox1, Refresh_Data
	// cannot be a control event and produces no event-handler finding.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		formFile("code/Dialog.bas", "Option Explicit\nPublic Sub Refresh_Data()\nEnd Sub\n"),
		formFile("Dialog.frm", "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog\n   Begin MSForms.TextBox TextBox1\n   End\nEnd\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA281"); len(got) != 0 {
		t.Fatalf("VBA281 findings = %+v, want none for a non-control form helper", got)
	}
}

func TestVBA281FlagsExplicitPublicControlEvent(t *testing.T) {
	// TextBox1 exists in the designer, so TextBox1_Change is a real control
	// event and an explicit Public declaration still leaks it.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		formFile("code/Dialog.bas", "Option Explicit\nPublic Sub TextBox1_Change()\nEnd Sub\n"),
		formFile("Dialog.frm", "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog\n   Begin MSForms.TextBox TextBox1\n   End\nEnd\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA281")
	if len(got) != 1 || !strings.Contains(got[0].Message, "TextBox1_Change") {
		t.Fatalf("VBA281 findings = %+v, want one finding for the Public control event", got)
	}
}

// --- VBA282: predeclared-instance self-name access ------------------------

func TestVBA282FlagsSelfNameAccessInDocumentModule(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate mCount As Long\nPublic Sub Bump()\n  Sheet1.mCount = Sheet1.mCount + 1\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	got := findingsByCode(findings, "VBA282")
	if len(got) == 0 {
		t.Fatalf("VBA282 findings = %+v, want findings for Sheet1 self-name access", got)
	}
}

func TestVBA282SkipsTypeAndDeclarationPositions(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Attribute VB_PredeclaredId = True\nOption Explicit\nPublic Sub Make()\n  Dim other As New Widget\n  Dim current As Widget\n  If TypeOf current Is Widget Then\n  End If\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	if got := findingsByCode(findings, "VBA282"); len(got) != 0 {
		t.Fatalf("VBA282 findings = %+v, want none for type/declaration positions", got)
	}
}

func TestVBA282FlagsTypeOfLeftOperandSelfName(t *testing.T) {
	// TypeOf Widget.Current Is Widget reads the default instance on the left;
	// only the trailing type operand may be excluded.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Attribute VB_PredeclaredId = True\nOption Explicit\nPublic Sub Probe()\n  If TypeOf Widget Is Widget Then\n  End If\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	got := findingsByCode(findings, "VBA282")
	if len(got) != 1 {
		t.Fatalf("VBA282 findings = %+v, want one finding for the TypeOf left operand", got)
	}
}

func TestVBA282SkipsLocalShadowingVariable(t *testing.T) {
	// A local variable named after the containing module shadows the class
	// name; the access resolves to ScopeLocal and must not warn.
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Attribute VB_PredeclaredId = True\nOption Explicit\nPublic Sub Make()\n  Dim Widget As Object\n  Set Widget = Me\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	if got := findingsByCode(findings, "VBA282"); len(got) != 0 {
		t.Fatalf("VBA282 findings = %+v, want none for a shadowing local variable", got)
	}
}

func TestVBA282RequiresPredeclaredId(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPrivate mCount As Long\nPublic Sub Bump()\n  Widget.mCount = Widget.mCount + 1\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	if got := findingsByCode(findings, "VBA282"); len(got) != 0 {
		t.Fatalf("VBA282 findings = %+v, want none without VB_PredeclaredId=True", got)
	}
}

func TestVBA282FlagsPredeclaredClassAttribute(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Attribute VB_PredeclaredId = True\nOption Explicit\nPublic Sub Bump()\n  Widget.Refresh\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	got := findingsByCode(findings, "VBA282")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Widget") {
		t.Fatalf("VBA282 findings = %+v, want one finding for the self-name call", got)
	}
}

// --- default-off ----------------------------------------------------------

func TestClassHazardsDisabledByDefault(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Sub Do_Work()\nEnd Sub\nPublic Property Let Size(v As Long)\nEnd Property\n"),
		documentFile("Sheet1.cls", "Option Explicit\nPublic Enum SheetMode\n  SheetRead\nEnd Enum\n"),
	}, nil)
	for _, code := range []string{"VBA278", "VBA279", "VBA280", "VBA281", "VBA282"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("%s findings = %+v, want none when the rule is disabled", code, got)
		}
	}
}
