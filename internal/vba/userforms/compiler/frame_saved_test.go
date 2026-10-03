package compiler

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// COM Controls enumeration does not represent persisted sibling order. Bind
// that contract to the binary saved after successful runtime execution instead.
func TestExcelSavedTemplateFrameHierarchy(t *testing.T) {
	dir := filepath.Join("testdata", "frame-excel-generated")
	data, err := os.ReadFile(filepath.Join(dir, "cli-template.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(data)
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Forms) != 1 {
		t.Fatal("expected one form")
	}
	actual, err := projection.Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	data, err = os.ReadFile(filepath.Join(dir, "cli-template-expected.json"))
	if err != nil {
		t.Fatal(err)
	}
	var expected spec.FormSpec
	if err := json.Unmarshal(data, &expected); err != nil {
		t.Fatal(err)
	}
	parents := func(value spec.FormSpec) map[string]string {
		result := map[string]string{}
		for _, control := range value.Controls {
			result[control.ID] = control.Name
		}
		return result
	}
	wantParents, gotParents := parents(expected), parents(actual)
	if len(expected.Controls) != len(actual.Controls) {
		t.Fatal("saved controls differ")
	}
	byName := map[string]spec.FormSpecControl{}
	for _, control := range actual.Controls {
		byName[control.Name] = control
	}
	for _, want := range expected.Controls {
		got, ok := byName[want.Name]
		if !ok || got.Type != want.Type || got.ProgID != want.ProgID || gotParents[got.ParentID] != wantParents[want.ParentID] || got.ZIndex == nil || want.ZIndex == nil || *got.ZIndex != *want.ZIndex {
			t.Errorf("saved hierarchy for %s: got parent=%s zIndex=%v; want parent=%s zIndex=%v", want.Name, gotParents[got.ParentID], got.ZIndex, wantParents[want.ParentID], want.ZIndex)
		}
	}
}
