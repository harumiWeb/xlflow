package pack

import (
	"errors"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
)

func TestPlanProjectIsReadOnlyAndRecordsDeterministicAuthority(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	before := append([]vbaproject.Module(nil), project.Modules...)
	sources := []SourceModule{
		{SourcePath: "src/modules/Zed.bas", Name: "Zed", Type: ModuleTypeStandard, Source: standardSource("Zed")},
		{Name: "ThisWorkbook", Type: ModuleTypeDocument, SourcePath: "src/workbook/ThisWorkbook.bas", Source: "Option Explicit\r\n"},
		{SourcePath: "src/classes/Alpha.cls", Name: "Alpha", Type: ModuleTypeClass, Source: classSource(t, "Alpha")},
	}
	plan, err := PlanProject(project, sources)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(project.Modules, before) {
		t.Fatal("planner modified template project")
	}
	var additions []string
	for _, component := range plan.Components {
		if component.Action == PlanAdd {
			additions = append(additions, component.Name)
			if component.TopologyAuthority != AuthoritySource || component.CodeAuthority != AuthoritySource {
				t.Fatalf("source-owned component authority = %+v", component)
			}
		}
		if component.Name == "ThisWorkbook" && (component.TopologyAuthority != AuthorityTemplate || component.CodeAuthority != AuthoritySource) {
			t.Fatalf("document authority = %+v", component)
		}
	}
	if !reflect.DeepEqual(additions, []string{"Alpha", "Zed"}) {
		t.Fatalf("additions = %#v", additions)
	}
}

func TestPlanProjectRejectsComprehensiveSourceIdentityFailures(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	cases := []struct {
		name    string
		sources []SourceModule
	}{
		{"case-insensitive duplicate", []SourceModule{{Name: "Thing", Type: ModuleTypeStandard, Source: standardSource("Thing")}, {Name: "thing", Type: ModuleTypeClass, Source: classSource(t, "thing")}}},
		{"invalid name", []SourceModule{{Name: "1Bad", Type: ModuleTypeStandard, Source: standardSource("1Bad")}}},
		{"unrepresentable writer name", []SourceModule{{Name: "Ābc", Type: ModuleTypeStandard, Source: standardSource("Ābc")}}},
		{"filename identity mismatch", []SourceModule{{SourcePath: "src/modules/New.bas", Name: "New", Type: ModuleTypeStandard, Source: standardSource("Old")}}},
		{"template-owned document case mismatch", []SourceModule{{Name: "thisworkbook", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"}}},
		{"template-owned document type mismatch", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeForm, Source: "Attribute VB_Name = \"ThisWorkbook\"\r\n"}}},
		{"document attribute header", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Attribute VB_Name = \"ThisWorkbook\"\r\nOption Explicit\r\n"}}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			_, err := PlanProject(project, tc.sources)
			if !errors.Is(err, ErrAmbiguousLayout) {
				t.Fatalf("err = %v, want ErrAmbiguousLayout", err)
			}
		})
	}
}

func TestPlanProjectSourceOwnedReplacementAllowsCaseAndTypeChange(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// Source-owned standard/class components need not preserve template case or
	// type: a case-only rename and a class-to-standard conversion plan as
	// removal plus addition rather than a rejection.
	sources := []SourceModule{
		{Name: "module1", Type: ModuleTypeStandard, Source: standardSource("module1")},
		{Name: "Class1", Type: ModuleTypeStandard, Source: standardSource("Class1")},
		{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"},
		{Name: "Sheet1", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"},
	}
	plan, err := PlanProject(project, sources)
	if err != nil {
		t.Fatal(err)
	}
	actions := map[string]PlanAction{}
	for _, component := range plan.Components {
		actions[string(component.Type)+":"+component.Name] = component.Action
	}
	want := map[string]PlanAction{
		"standard:Module1":      PlanRemove,
		"class:Class1":          PlanRemove,
		"standard:module1":      PlanAdd,
		"standard:Class1":       PlanAdd,
		"document:ThisWorkbook": PlanUpdate,
		"document:Sheet1":       PlanUpdate,
	}
	if !reflect.DeepEqual(actions, want) {
		t.Fatalf("actions = %#v, want %#v", actions, want)
	}
}

func TestPlanProjectAcceptsCodepageRepresentableNames(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// CP932 can represent Japanese component names; planning must allow them
	// for source-owned standard/class modules.
	sources := []SourceModule{
		{Name: "標準モジュール", Type: ModuleTypeStandard, Source: standardSource("標準モジュール")},
		{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"},
		{Name: "Sheet1", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"},
	}
	plan, err := PlanProject(project, sources)
	if err != nil {
		t.Fatalf("PlanProject: %v", err)
	}
	var found bool
	for _, component := range plan.Components {
		if component.Name == "標準モジュール" && component.Action == PlanAdd {
			found = true
		}
	}
	if !found {
		t.Fatalf("Japanese module not planned for addition: %#v", plan.Components)
	}
}

func TestPlanProjectRejectsSourceOnlyTemplateOwnedTopology(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := PlanProject(project, []SourceModule{{Name: "Sheet99", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"}}); !errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("document err = %v", err)
	}
	if _, err := PlanProject(project, []SourceModule{{Name: "Form99", Type: ModuleTypeForm, Source: "Attribute VB_Name = \"Form99\"\r\n"}}); err == nil {
		t.Fatalf("form err = %v", err)
	} else if detail, ok := errors.AsType[*compiler.Error](err); !ok || detail.Code != compiler.GenerationUnsupported {
		t.Fatalf("wrong capability error = %v", err)
	}
}

func TestPlanProjectRejectsCFBEquivalentStreamNames(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// Upcased UTF-16 code units are the CFB directory equivalence, so dotted
	// "I" and dotless "\u0131" collide even though they are distinct lowercase
	// keys. CP1254 can represent both.
	project.Props.CodePage = 1254
	sources := []SourceModule{
		{Name: "I", Type: ModuleTypeStandard, Source: standardSource("I")},
		{Name: "\u0131", Type: ModuleTypeStandard, Source: standardSource("\u0131")},
	}
	if _, err := PlanProject(project, sources); !errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("err = %v, want ErrAmbiguousLayout", err)
	}
}
