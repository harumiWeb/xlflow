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

// --- VBA273: public member underscore names -----------------------------

func TestVBA273FlagsPublicUnderscoredMemberInClass(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Sub Do_Work()\nEnd Sub\nPublic Sub Run()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	got := findingsByCode(findings, "VBA273")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Do_Work") {
		t.Fatalf("VBA273 findings = %+v, want one finding for Do_Work", got)
	}
	if got[0].Line != 2 {
		t.Fatalf("VBA273 line = %d, want the Sub declaration line 2", got[0].Line)
	}
}

func TestVBA273SkipsPrivateAndFriendMembers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPrivate Sub Helper_One()\nEnd Sub\nFriend Sub Helper_Two()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	if got := findingsByCode(findings, "VBA273"); len(got) != 0 {
		t.Fatalf("VBA273 findings = %+v, want none for private/friend members", got)
	}
}

func TestVBA273SkipsStandardModules(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		standardFile("Main.bas", "Option Explicit\nPublic Sub Do_Work()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	if got := findingsByCode(findings, "VBA273"); len(got) != 0 {
		t.Fatalf("VBA273 findings = %+v, want none in standard modules", got)
	}
}

func TestVBA273SkipsEventHandlers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\nPublic Sub Worksheet_Follow()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	for _, finding := range findingsByCode(findings, "VBA273") {
		if strings.Contains(finding.Message, "Worksheet_Change") {
			t.Fatalf("VBA273 flagged the event handler: %+v", finding)
		}
	}
}

func TestVBA273SkipsImplementedInterfaceMembers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPrivate Sub IFace_Go()\nEnd Sub\nPublic Sub IFace_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	for _, finding := range findingsByCode(findings, "VBA273") {
		if strings.Contains(finding.Message, "IFace_Go") {
			t.Fatalf("VBA273 flagged an implemented-interface member: %+v", finding)
		}
	}
}

func TestVBA273SkipsClassLifecycleMembers(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Sub Class_Initialize()\nEnd Sub\nPublic Sub Class_Terminate()\nEnd Sub\nPublic Sub Other_Name()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicMemberUnderscoreNames = true })
	got := findingsByCode(findings, "VBA273")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Other_Name") {
		t.Fatalf("VBA273 findings = %+v, want a single Other_Name finding", got)
	}
}

// --- VBA274: public Enum in document modules -----------------------------

func TestVBA274FlagsPublicEnumInDocumentModule(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPublic Enum SheetMode\n  SheetRead\n  SheetWrite\nEnd Enum\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectDocumentModulePublicEnum = true })
	got := findingsByCode(findings, "VBA274")
	if len(got) != 1 || !strings.Contains(got[0].Message, "SheetMode") {
		t.Fatalf("VBA274 findings = %+v, want one finding for SheetMode", got)
	}
	if got[0].Line != 2 {
		t.Fatalf("VBA274 line = %d, want the Enum declaration line 2", got[0].Line)
	}
}

func TestVBA274SkipsPrivateEnumAndOtherModuleKinds(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate Enum SheetMode\n  SheetRead\nEnd Enum\n"),
		standardFile("Main.bas", "Option Explicit\nPublic Enum Mode\n  ModeRead\nEnd Enum\n"),
		classFile("Widget.cls", "Option Explicit\nPublic Enum WidgetMode\n  WidgetRead\nEnd Enum\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectDocumentModulePublicEnum = true })
	if got := findingsByCode(findings, "VBA274"); len(got) != 0 {
		t.Fatalf("VBA274 findings = %+v, want none for private/other-module enums", got)
	}
}

// --- VBA275: write-only property ----------------------------------------

func TestVBA275FlagsWriteOnlyProperty(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Size(v As Long)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	got := findingsByCode(findings, "VBA275")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Size") || !strings.Contains(got[0].Message, "write-only") {
		t.Fatalf("VBA275 findings = %+v, want one write-only finding for Size", got)
	}
}

func TestVBA275SatisfiedByMatchingGet(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Size(v As Long)\nEnd Property\nPublic Property Get Size() As Long\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none when a Get exists", got)
	}
}

func TestVBA275SkipsPrivateWriters(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPrivate Property Let Size(v As Long)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for private writers", got)
	}
}

func TestVBA275DeduplicatesLetAndSet(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Property Let Value(v As Long)\nEnd Property\nPublic Property Set Value(v As Object)\nEnd Property\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectWriteOnlyProperty = true })
	if got := findingsByCode(findings, "VBA275"); len(got) != 1 {
		t.Fatalf("VBA275 findings = %+v, want one deduplicated finding", got)
	}
}

// --- VBA276: interface/event exposure ------------------------------------

func TestVBA276FlagsPublicImplementedInterfaceMember(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPublic Sub IFace_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA276")
	if len(got) != 1 || !strings.Contains(got[0].Message, "IFace_Go") {
		t.Fatalf("VBA276 findings = %+v, want one finding for IFace_Go", got)
	}
}

func TestVBA276MatchesFullUnderscoredInterfaceName(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nImplements I_Foo\nPublic Sub I_Foo_Bar()\nEnd Sub\nPublic Sub I_Other()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA276")
	if len(got) != 1 || !strings.Contains(got[0].Message, "I_Foo_Bar") {
		t.Fatalf("VBA276 findings = %+v, want only I_Foo_Bar", got)
	}
}

func TestVBA276FlagsExplicitPublicEventHandler(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPublic Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	got := findingsByCode(findings, "VBA276")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Worksheet_Change") {
		t.Fatalf("VBA276 findings = %+v, want one finding for the explicit Public event handler", got)
	}
}

func TestVBA276SkipsImplicitVisibilityEventHandler(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\nSub Worksheet_SelectionChange(ByVal Target As Range)\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA276"); len(got) != 0 {
		t.Fatalf("VBA276 findings = %+v, want none without explicit Public", got)
	}
}

func TestVBA276SkipsPrivateImplementations(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nImplements IFace\nPrivate Sub IFace_Go()\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPublicInterfaceEventMembers = true })
	if got := findingsByCode(findings, "VBA276"); len(got) != 0 {
		t.Fatalf("VBA276 findings = %+v, want none for private implementations", got)
	}
}

// --- VBA277: predeclared-instance self-name access ------------------------

func TestVBA277FlagsSelfNameAccessInDocumentModule(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		documentFile("Sheet1.cls", "Option Explicit\nPrivate mCount As Long\nPublic Sub Bump()\n  Sheet1.mCount = Sheet1.mCount + 1\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	got := findingsByCode(findings, "VBA277")
	if len(got) == 0 {
		t.Fatalf("VBA277 findings = %+v, want findings for Sheet1 self-name access", got)
	}
}

func TestVBA277SkipsTypeAndDeclarationPositions(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Attribute VB_PredeclaredId = True\nOption Explicit\nPublic Sub Make()\n  Dim other As New Widget\n  Dim current As Widget\n  If TypeOf current Is Widget Then\n  End If\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	if got := findingsByCode(findings, "VBA277"); len(got) != 0 {
		t.Fatalf("VBA277 findings = %+v, want none for type/declaration positions", got)
	}
}

func TestVBA277RequiresPredeclaredId(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPrivate mCount As Long\nPublic Sub Bump()\n  Widget.mCount = Widget.mCount + 1\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	if got := findingsByCode(findings, "VBA277"); len(got) != 0 {
		t.Fatalf("VBA277 findings = %+v, want none without VB_PredeclaredId=True", got)
	}
}

func TestVBA277FlagsPredeclaredClassAttribute(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Attribute VB_PredeclaredId = True\nOption Explicit\nPublic Sub Bump()\n  Widget.Refresh\nEnd Sub\n"),
	}, func(cfg *config.AnalyzeConfig) { cfg.DetectPredeclaredInstanceAccess = true })
	got := findingsByCode(findings, "VBA277")
	if len(got) != 1 || !strings.Contains(got[0].Message, "Widget") {
		t.Fatalf("VBA277 findings = %+v, want one finding for the self-name call", got)
	}
}

// --- default-off ----------------------------------------------------------

func TestClassHazardsDisabledByDefault(t *testing.T) {
	findings := runClassHazardProjectAnalysis(t, []sourceproject.SourceFile{
		classFile("Widget.cls", "Option Explicit\nPublic Sub Do_Work()\nEnd Sub\nPublic Property Let Size(v As Long)\nEnd Property\n"),
		documentFile("Sheet1.cls", "Option Explicit\nPublic Enum SheetMode\n  SheetRead\nEnd Enum\n"),
	}, nil)
	for _, code := range []string{"VBA273", "VBA274", "VBA275", "VBA276", "VBA277"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("%s findings = %+v, want none when the rule is disabled", code, got)
		}
	}
}
