package pack

import (
	"errors"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
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
		{"filename identity mismatch", []SourceModule{{SourcePath: "src/modules/New.bas", Name: "New", Type: ModuleTypeStandard, Source: standardSource("Old")}}},
		{"template-owned document case mismatch", []SourceModule{{Name: "thisworkbook", Type: ModuleTypeDocument, Source: "Option Explicit\r\n"}}},
		{"template-owned document type mismatch", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeForm, Source: "Attribute VB_Name = \"ThisWorkbook\"\r\n"}}},
		{"document attribute header", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Attribute VB_Name = \"ThisWorkbook\"\r\nOption Explicit\r\n"}}},
		{"document host attribute header", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"\r\nOption Explicit\r\n"}}},
		{"case-insensitive document host attribute header", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"\r\nOption Explicit\r\n"}}},
		{"unknown document member attribute", []SourceModule{{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Option Explicit\r\nAttribute cbs.VB_Name = \"forged\"\r\n"}}},
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

func TestPlanProjectPreservesDocumentMemberAttributes(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	const memberAttribute = "Attribute cbs.VB_VarHelpID = -1"
	savedBody := "Option Explicit\n" + memberAttribute + "\n" +
		"Private Sub OriginalCode()\nEnd Sub\n"
	seedDocumentModuleBody(t, project, "Sheet1", savedBody)
	source := "Option Explicit\n" + memberAttribute + "\n" +
		"Private Sub EditedCode()\nEnd Sub\n"
	plan, err := PlanProject(project, []SourceModule{{
		Name:       "Sheet1",
		Type:       ModuleTypeDocument,
		SourcePath: "src/workbook/Sheet1.bas",
		Source:     source,
	}})
	if err != nil {
		t.Fatal(err)
	}
	for _, module := range plan.modules {
		if module.Name != "Sheet1" {
			continue
		}
		if !strings.Contains(module.Source, "Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\r\n") {
			t.Fatalf("template host attributes were not retained: %q", module.Source)
		}
		if !strings.Contains(module.Source, memberAttribute+"\r\n") || !strings.Contains(module.Source, "Private Sub EditedCode()\r\n") {
			t.Fatalf("saved member attribute or edited code was not retained: %q", module.Source)
		}
		return
	}
	t.Fatal("plan omitted template document module Sheet1")
}

func TestPlanProjectRejectsNewOrModifiedDocumentMemberAttributes(t *testing.T) {
	const savedAttribute = "Attribute cbs.VB_VarHelpID = -1"
	cases := []struct {
		name   string
		source string
	}{
		{name: "new member attribute", source: "Option Explicit\n" + savedAttribute + "\nAttribute cbs.VB_Description = \"new\"\n"},
		{name: "modified saved attribute", source: "Option Explicit\nAttribute cbs.VB_VarHelpID = 1\n"},
		{name: "duplicate saved attribute", source: "Option Explicit\n" + savedAttribute + "\n" + savedAttribute + "\n"},
		{name: "whitespace modified saved statement", source: "Option Explicit\n" + savedAttribute + " \n"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
			if err != nil {
				t.Fatal(err)
			}
			seedDocumentModuleBody(t, project, "Sheet1", "Option Explicit\n"+savedAttribute+"\n")
			_, err = PlanProject(project, []SourceModule{{
				Name:       "Sheet1",
				Type:       ModuleTypeDocument,
				SourcePath: "src/workbook/Sheet1.bas",
				Source:     tc.source,
			}})
			if !errors.Is(err, ErrAmbiguousLayout) {
				t.Fatalf("err = %v, want ErrAmbiguousLayout", err)
			}
		})
	}
}

func seedDocumentModuleBody(t *testing.T, project *vbaproject.Project, name, body string) {
	t.Helper()
	for index := range project.Modules {
		module := &project.Modules[index]
		if module.Name != name {
			continue
		}
		var header []string
		for _, line := range strings.Split(strings.ReplaceAll(module.Source, "\r\n", "\n"), "\n") {
			if !strings.HasPrefix(line, "Attribute ") {
				break
			}
			header = append(header, line)
		}
		if len(header) == 0 {
			t.Fatalf("document module %s has no leading host attributes", name)
		}
		module.Source = strings.Join(header, "\r\n") + "\r\n" + strings.ReplaceAll(body, "\n", "\r\n")
		return
	}
	t.Fatalf("template has no document module %s", name)
}

func TestPlanProjectReusesExportEquivalentDocumentSource(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	var sheet *vbaproject.Module
	for index := range project.Modules {
		if project.Modules[index].Name == "Sheet1" {
			sheet = &project.Modules[index]
			break
		}
	}
	if sheet == nil {
		t.Fatal("template has no Sheet1 document module")
		return
	}
	sheet.Source = strings.TrimRight(sheet.Source, "\r\n") + "\r\n\r\n"
	disk, err := vbaproject.ExportModuleSource(*sheet)
	if err != nil {
		t.Fatal(err)
	}
	plan, err := PlanProject(project, []SourceModule{{
		Name:       sheet.Name,
		Type:       ModuleTypeDocument,
		SourcePath: "src/workbook/Sheet1.bas",
		Source:     disk,
	}})
	if err != nil {
		t.Fatal(err)
	}
	for _, module := range plan.modules {
		if module.Name == sheet.Name {
			if module.Source != sheet.Source {
				t.Fatalf("no-op plan changed document source\n got: %q\nwant: %q", module.Source, sheet.Source)
			}
			return
		}
	}
	t.Fatal("plan omitted template document module Sheet1")
}

func TestPlanProjectReportsTypedSourceEncodingLocations(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project.Props.CodePage = 1252
	sources := []SourceModule{
		{SourcePath: "src/modules/Emoji.bas", Name: "Emoji", Type: ModuleTypeStandard, Source: standardSource("Emoji") + "' \U0001F600\r\n"},
		{SourcePath: "src/modules/Japanese.bas", Name: "Japanese", Type: ModuleTypeStandard, Source: standardSource("Japanese") + "' 日本語\r\n"},
	}
	_, err = PlanProject(project, sources)
	if errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("source encoding was misclassified as ambiguous layout: %v", err)
	}
	diagnostics, ok := errors.AsType[*SourceEncodingErrors](err)
	if !ok || len(diagnostics.Errors) != 2 {
		t.Fatalf("encoding diagnostics = %#v, want two source errors", err)
	}
	want := map[string]rune{"src/modules/Emoji.bas": '\U0001F600', "src/modules/Japanese.bas": '日'}
	for _, issue := range diagnostics.Errors {
		if got := want[issue.Path]; got == 0 || issue.Rune != got || issue.CodePage != 1252 {
			t.Fatalf("unexpected issue: %+v", issue)
		}
		if issue.Line != 3 || issue.Column != 3 {
			t.Errorf("%s position = %d:%d, want 3:3", issue.Path, issue.Line, issue.Column)
		}
		delete(want, issue.Path)
	}
	if len(want) != 0 {
		t.Fatalf("missing source issues: %v", want)
	}
}

func TestPlanProjectDoesNotClassifyUnsupportedCodePageAsAmbiguousLayout(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project.Props.CodePage = 37
	_, err = PlanProject(project, []SourceModule{{Name: "Module1", Type: ModuleTypeStandard, Source: standardSource("Module1")}})
	if errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("unsupported codepage was misclassified as ambiguous layout: %v", err)
	}
	codePageErr, ok := errors.AsType[*ovba.CodePageEncodingError](err)
	if !ok || codePageErr.CodePage != 37 || !codePageErr.UnsupportedCodePage {
		t.Fatalf("codepage error = %#v, want typed unsupported-codepage error", err)
	}
}

func TestPlanProjectAggregatesIndependentFormCompilerErrors(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	first := authoredForm("FirstBad")
	first.FormSpec.Controls[0].Type = "UnknownControl"
	second := authoredForm("SecondBad")
	second.FormSpec.Controls[0].Width = new(-1.0)
	_, err = PlanProject(project, []SourceModule{first, second})
	diagnostics, ok := errors.AsType[*DiagnosticsError](err)
	if !ok || len(diagnostics.Errors) != 2 {
		t.Fatalf("diagnostics = %#v, want two compiler errors", err)
	}
	for _, issue := range diagnostics.Errors {
		if _, ok := errors.AsType[*compiler.Error](issue); !ok {
			t.Errorf("diagnostic %T = %v, want compiler error", issue, issue)
		}
	}
}

func TestPlanProjectEncodingAndFormCompilerErrorsAreCombined(t *testing.T) {
	project, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project.Props.CodePage = 1252
	badForm := authoredForm("BadForm")
	badForm.FormSpec.Form.Caption = new("Bad form")
	badForm.FormSpec.Controls[0].Type = "UnknownControl"
	sources := []SourceModule{
		{SourcePath: "src/modules/Bad.bas", Name: "Bad", Type: ModuleTypeStandard, Source: standardSource("Bad") + "' \U0001F600\r\n"},
		badForm,
	}
	_, err = PlanProject(project, sources)
	diagnostics, ok := errors.AsType[*DiagnosticsError](err)
	if !ok || len(diagnostics.Errors) != 2 {
		t.Fatalf("diagnostics = %#v, want encoding and compiler issues", err)
	}
	if _, ok := errors.AsType[*SourceEncodingErrors](err); !ok {
		t.Fatalf("combined error lost source encoding issues: %v", err)
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
