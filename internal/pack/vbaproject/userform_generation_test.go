package vbaproject

import (
	"bytes"
	"encoding/binary"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/ovba"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
)

func TestNewUserFormModuleUsesInjectedGUIDs(t *testing.T) {
	form := generatedTestForm(t, "GenerationBaseline", 932)
	guids := []string{
		"11111111-2222-4333-8444-555555555555",
		"AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE",
	}
	index := 0
	module, err := NewUserFormModule(form, "Option Explicit\n", 932, UserFormModuleOptions{
		GUIDGenerator: func() (string, error) {
			value := guids[index]
			index++
			return value, nil
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	want := strings.Join([]string{
		`Attribute VB_Name = "GenerationBaseline"`,
		`Attribute VB_Base = "0{11111111-2222-4333-8444-555555555555}{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}"`,
		"Attribute VB_GlobalNameSpace = False",
		"Attribute VB_Creatable = False",
		"Attribute VB_PredeclaredId = True",
		"Attribute VB_Exposed = False",
		"Attribute VB_TemplateDerived = False",
		"Attribute VB_Customizable = False",
		"Option Explicit",
	}, "\r\n") + "\r\n"
	if module.Type != ModuleForm || module.Name != form.Name || module.StreamName != form.Name {
		t.Fatalf("module identity = %+v", module)
	}
	if module.Source != want {
		t.Fatalf("module source mismatch\n got=%q\nwant=%q", module.Source, want)
	}
	if !strings.Contains(form.DesignerSource.Text, "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GenerationBaseline") {
		t.Fatal("generated Designer does not retain the fixed MS-OFORMS VBFrame identity")
	}
}

func TestNewUserFormModuleConfirmsExcelAuthoredComponentIdentity(t *testing.T) {
	body := readUserFormGenerationFixture(t, "00-baseline.bin")
	project, err := Read(body)
	if err != nil {
		t.Fatal(err)
	}
	var found *Module
	for i := range project.Modules {
		if project.Modules[i].Type == ModuleForm {
			found = &project.Modules[i]
			break
		}
	}
	if found == nil {
		t.Fatal("Excel-authored baseline has no UserForm module")
	}
	if err := validateUserFormModuleSource(*found, found.Name); err != nil {
		t.Fatalf("baseline UserForm module identity: %v", err)
	}
	base := strings.Split(found.Source, "\r\n")[1]
	if strings.Count(base, "{") != 2 || strings.Count(base, "}") != 2 {
		t.Fatalf("baseline VB_Base does not contain two component GUIDs: %q", base)
	}
}

func TestWithNewUserFormRoundTripsAndPreservesTemplateState(t *testing.T) {
	templateBytes := loadCorpus(t, "p4_form")
	template, err := Read(templateBytes)
	if err != nil {
		t.Fatal(err)
	}
	before := cloneProject(template)
	form := generatedTestForm(t, "GeneratedForm", template.Props.CodePage)
	module, err := NewUserFormModule(form, "Option Explicit\r\n", template.Props.CodePage, UserFormModuleOptions{
		GUIDGenerator: fixedGUIDGenerator(),
	})
	if err != nil {
		t.Fatal(err)
	}
	result, err := WithNewUserForm(template, form, module)
	if err != nil {
		t.Fatal(err)
	}
	if result == template || result.Forms[len(result.Forms)-1] == form {
		t.Fatal("WithNewUserForm did not return independent project/form state")
	}
	if !reflect.DeepEqual(template, before) {
		t.Fatal("WithNewUserForm mutated its input project")
	}

	written, err := Write(result)
	if err != nil {
		t.Fatal(err)
	}
	reparsed, err := Read(written)
	if err != nil {
		t.Fatalf("pure-Go reparse: %v", err)
	}
	if !hasModule(reparsed, "GeneratedForm", ModuleForm) || !hasForm(reparsed, "GeneratedForm") {
		t.Fatalf("reparsed project lost generated form topology: modules=%v forms=%v", reparsed.Modules, reparsed.Forms)
	}
	if !bytes.Equal(reparsed.ReferencesRaw, template.ReferencesRaw) {
		t.Fatal("existing references changed")
	}

	originalCFB, err := cfb.Open(templateBytes)
	if err != nil {
		t.Fatal(err)
	}
	resultCFB, err := cfb.Open(written)
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range originalCFB.Paths() {
		if !strings.HasPrefix(path, "UserForm1/") {
			continue
		}
		want, _ := originalCFB.Stream(path)
		have, ok := resultCFB.Stream(path)
		if !ok || !bytes.Equal(have, want) {
			t.Fatalf("existing Designer stream %q changed", path)
		}
	}
	expectedProject, err := addUserFormProjectComponent(template.ProjectStreamRaw, "GeneratedForm", template.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	actualProject, ok := resultCFB.Stream("PROJECT")
	if !ok || !bytes.Equal(actualProject, expectedProject) {
		t.Fatal("PROJECT did not preserve template text while adding BaseClass")
	}
	if !bytes.Contains(actualProject, []byte("BaseClass=GeneratedForm\r\n")) {
		t.Fatal("PROJECT is missing the generated BaseClass declaration")
	}

	var specs []ovba.ProjectComponentSpec
	for _, item := range result.Modules {
		specs = append(specs, ovba.ProjectComponentSpec{Kind: projectKind(item.Type), Name: item.Name})
	}
	expectedWM, err := ovba.BuildProjectWM(specs, result.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	actualWM, ok := resultCFB.Stream("PROJECTwm")
	if !ok || !bytes.Equal(actualWM, expectedWM) {
		t.Fatal("PROJECTwm was not rebuilt for the added template component")
	}
}

func TestWithNewUserFormRejectsWithoutMutation(t *testing.T) {
	tests := []struct {
		name  string
		book  string
		setup func(*Project, *oforms.Form, *Module)
		code  string
	}{
		{
			name: "forms reference required",
			book: "p2_refs",
			setup: func(project *Project, _ *oforms.Form, _ *Module) {
				project.References = append(project.References, Reference{Name: "MSForms"})
			},
			code: UserFormFormsReferenceRequired,
		},
		{
			name: "protected project",
			book: "p3_protected",
			code: UserFormGenerationInvalid,
		},
		{
			name: "module attribute mismatch",
			book: "p4_form",
			setup: func(_ *Project, _ *oforms.Form, module *Module) {
				module.Source = strings.Replace(module.Source, "Attribute VB_Exposed = False", "Attribute VB_Exposed = True", 1)
			},
			code: UserFormGenerationInvalid,
		},
		{
			name: "module code page mismatch",
			book: "p4_form",
			setup: func(_ *Project, _ *oforms.Form, module *Module) {
				module.Source += "😀\r\n"
			},
			code: UserFormGenerationInvalid,
		},
		{
			name: "module name collision",
			book: "p4_form",
			setup: func(project *Project, form *oforms.Form, module *Module) {
				form.Name = "UserForm1"
				module.Name = form.Name
				module.StreamName = form.Name
				module.Source = strings.Replace(module.Source, `VB_Name = "GeneratedForm"`, `VB_Name = "UserForm1"`, 1)
				project.Forms = project.Forms[:1]
			},
			code: UserFormGenerationConflict,
		},
		{
			name: "reserved top storage",
			book: "p4_form",
			setup: func(project *Project, _ *oforms.Form, _ *Module) {
				project.RawStreams["GeneratedForm"] = []byte("occupied")
			},
			code: UserFormGenerationConflict,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			project, err := Read(loadCorpus(t, test.book))
			if err != nil {
				t.Fatal(err)
			}
			form := generatedTestForm(t, "GeneratedForm", project.Props.CodePage)
			module, err := NewUserFormModule(form, "Option Explicit", project.Props.CodePage, UserFormModuleOptions{GUIDGenerator: fixedGUIDGenerator()})
			if err != nil {
				t.Fatal(err)
			}
			if test.setup != nil {
				test.setup(project, form, &module)
			}
			before := cloneProject(project)
			_, err = WithNewUserForm(project, form, module)
			detail, ok := errors.AsType[*UserFormGenerationError](err)
			if !ok || detail.Code != test.code {
				t.Fatalf("error = %v, want code %q", err, test.code)
			}
			if !reflect.DeepEqual(project, before) {
				t.Fatal("rejected addition mutated its input project")
			}
		})
	}
}

func TestHasMSFormsReferenceUsesRawDirRecords(t *testing.T) {
	metadataOnly := &Project{References: []Reference{{Name: "MSForms"}}}
	if hasMSFormsReference(metadataOnly) {
		t.Fatal("display metadata without ReferencesRaw was accepted")
	}

	raw := appendReferenceRecord(nil, 0x0016, []byte("MSForms"))
	raw = appendReferenceRecord(raw, 0x002F, []byte("1\x00\x00\x00*\\G{00000000-0000-0000-0000-000000000000}"))
	if !hasMSFormsReferenceRaw(raw) {
		t.Fatal("REFERENCECONTROL Forms evidence was not accepted")
	}
}

func generatedTestForm(t *testing.T, name string, codePage uint16) *oforms.Form {
	t.Helper()
	form, err := oforms.NewForm(oforms.Definition{Name: name, Size: oforms.Size{Width: 2540, Height: 2540}}, codePage)
	if err != nil {
		t.Fatal(err)
	}
	return form
}

func fixedGUIDGenerator() UserFormGUIDGenerator {
	guids := []string{
		"11111111-2222-4333-8444-555555555555",
		"AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE",
	}
	index := 0
	return func() (string, error) {
		value := guids[index]
		index++
		return value, nil
	}
}

func appendReferenceRecord(raw []byte, id uint16, payload []byte) []byte {
	record := make([]byte, 6+len(payload))
	binary.LittleEndian.PutUint16(record, id)
	binary.LittleEndian.PutUint32(record[2:], uint32(len(payload)))
	copy(record[6:], payload)
	return append(raw, record...)
}

func readUserFormGenerationFixture(t *testing.T, name string) []byte {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "..", "vba", "userforms", "compiler", "testdata", "excel-authored", name))
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func hasModule(project *Project, name string, moduleType ModuleType) bool {
	for _, module := range project.Modules {
		if module.Name == name && module.Type == moduleType {
			return true
		}
	}
	return false
}

func hasForm(project *Project, name string) bool {
	for _, form := range project.Forms {
		if form != nil && form.Name == name {
			return true
		}
	}
	return false
}
