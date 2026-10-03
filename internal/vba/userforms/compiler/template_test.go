package compiler

import (
	"bytes"
	"errors"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func templateInput(t *testing.T, form *oforms.Form, snapshot spec.FormSpec) spec.FormSpec {
	t.Helper()
	input := spec.FormSpec{
		SchemaVersion:    1,
		Kind:             "xlflow.userform",
		Basis:            "designer",
		CoordinateSystem: "parent-relative",
		Form:             spec.FormSpecForm{Name: form.Name},
		Controls:         make([]spec.FormSpecControl, len(snapshot.Controls)),
	}
	ids := make(map[string]string, len(snapshot.Controls))
	for index, control := range snapshot.Controls {
		ids[control.Name] = "source-id-" + strings.ReplaceAll(control.Name, " ", "_")
		input.Controls[index] = spec.FormSpecControl{
			ID:     ids[control.Name],
			Name:   control.Name,
			Type:   control.Type,
			ProgID: control.ProgID,
		}
	}
	for index, control := range snapshot.Controls {
		if control.ParentID != "" {
			parentName := ""
			for _, candidate := range snapshot.Controls {
				if candidate.ID == control.ParentID {
					parentName = candidate.Name
					break
				}
			}
			input.Controls[index].ParentID = ids[parentName]
		}
	}
	return input
}

func TestCompileTemplateNoOpPreservesBytesAndIgnoresSnapshotMetadata(t *testing.T) {
	for _, fixtureName := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		t.Run(fixtureName, func(t *testing.T) {
			base, snapshot := fixture(t, fixtureName)
			desired := templateInput(t, base, snapshot)
			desired.Warnings = []spec.FormSpecWarning{{Code: "observed-only"}}
			desired.Form.Observed = &spec.FormSpecObservedForm{Width: new(999.0), ClientWidth: new(888.0)}
			for index := range desired.Controls {
				desired.Controls[index].Observed = &spec.FormSpecObservedControl{Caption: new("ignored")}
				desired.Controls[index].Unsupported = []string{"opaqueRecordTail"}
			}
			beforeInput := snapshotCopy(t, desired)
			result, err := CompileTemplate(base, desired, 932)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(beforeInput, desired) {
				t.Fatal("CompileTemplate mutated desired input")
			}
			original, err := oforms.SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			compiled, err := oforms.SerializeForm(result, 932)
			if err != nil {
				t.Fatal(err)
			}
			assertDesignerBytes(t, original, compiled)
		})
	}
}

func TestCompileTemplateEqualExplicitRootCaptionRepairsLegacyVBFrame(t *testing.T) {
	base, _ := fixture(t, "p4_form.bin")
	updated, err := oforms.ApplyEdits(base, []oforms.Edit{{Control: "", Property: "Caption", Value: "Canonical Caption"}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	legacy, err := oforms.SerializeForm(updated, 932)
	if err != nil {
		t.Fatal(err)
	}
	// Older generated projects contain the requested Caption only in f.
	legacy.Streams["UserForm1/\x03VBFrame"] = []byte("VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\r\nEnd\r\n")
	writer := cfb.NewWriter()
	for path, metadata := range legacy.Storages {
		writer.AddStorage(strings.Split(path, "/"), metadata)
	}
	for path, raw := range legacy.Streams {
		writer.AddStream(strings.Split(path, "/"), raw)
	}
	body, err := writer.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	base, err = oforms.ReadForm(container, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	snapshot, err := projection.Project(base)
	if err != nil {
		t.Fatal(err)
	}
	desired := templateInput(t, base, snapshot)
	desired.Form.Caption = new(*snapshot.Form.Caption)

	before, err := oforms.SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	result, err := CompileTemplate(base, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	after, err := oforms.SerializeForm(result, 932)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Equal(before.Streams["UserForm1/\x03VBFrame"], after.Streams["UserForm1/\x03VBFrame"]) {
		t.Fatal("explicit equal root caption did not repair missing VBFrame Caption")
	}
	if !bytes.Contains(after.Streams["UserForm1/\x03VBFrame"], []byte(`Caption = "`+*snapshot.Form.Caption+`"`)) {
		t.Fatal("repaired VBFrame Caption is missing")
	}

	repeated, err := CompileTemplate(result, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	repeatedBytes, err := oforms.SerializeForm(repeated, 932)
	if err != nil {
		t.Fatal(err)
	}
	assertDesignerBytes(t, after, repeatedBytes)
}

func TestCompileTemplateOverlaysOmittedFieldsAndRemapsParents(t *testing.T) {
	base, snapshot := fixture(t, "p6_nested_form.bin")
	desired := templateInput(t, base, snapshot)
	if len(desired.Controls) < 2 {
		t.Fatal("nested fixture has too few controls")
	}
	desired.Controls[0].Caption = new("template caption")
	desired.Controls[0].Properties = map[string]any{"Caption": "template caption"}
	desired.Controls[1].Observed = &spec.FormSpecObservedControl{Caption: new("must remain ignored")}
	result, err := CompileTemplate(base, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	projected, err := projection.Project(result)
	if err != nil {
		t.Fatal(err)
	}
	if projected.Controls[0].Caption == nil || *projected.Controls[0].Caption != "template caption" {
		t.Fatalf("updated caption = %#v", projected.Controls[0].Caption)
	}
	if !reflect.DeepEqual(projected.Controls[1].Caption, snapshot.Controls[1].Caption) {
		t.Fatalf("omitted caption changed: got %#v want %#v", projected.Controls[1].Caption, snapshot.Controls[1].Caption)
	}
}

func TestCompileTemplateAcceptsEqualDimensionsAndRejectsChanges(t *testing.T) {
	base, snapshot := fixture(t, "p4_form.bin")
	desired := templateInput(t, base, snapshot)
	desired.Form.Build = &spec.FormSpecBuildForm{}
	if snapshot.Form.Build != nil {
		desired.Form.Build.ClientWidth = new(*snapshot.Form.Build.ClientWidth)
		desired.Form.Build.ClientHeight = new(*snapshot.Form.Build.ClientHeight)
	}
	if _, err := CompileTemplate(base, desired, 932); err != nil {
		t.Fatalf("equal dimensions rejected: %v", err)
	}
	desired.Form.Build.ClientWidth = new(1.0)
	_, err := CompileTemplate(base, desired, 932)
	detail, ok := errors.AsType[*Error](err)
	if !ok || detail.Code != Unsupported {
		t.Fatalf("changed dimensions error = %v, want %s", err, Unsupported)
	}
}

func TestCompileTemplateRejectsUnsupportedEditsAndAllowsOpaqueNoOp(t *testing.T) {
	base, snapshot := fixture(t, "p6_nested_form.bin")
	desired := templateInput(t, base, snapshot)
	for index, control := range snapshot.Controls {
		if control.Type == "MultiPage" {
			continue
		}
		desired.Controls[index].Properties = map[string]any{"Picture": "unsupported"}
		_, err := CompileTemplate(base, desired, 932)
		detail, ok := errors.AsType[*Error](err)
		if !ok || detail.Code != Unsupported {
			t.Fatalf("unsupported property error = %v, want %s", err, Unsupported)
		}
		return
	}
	t.Fatal("nested fixture did not contain an editable control")
}

func TestCompileTemplateRejectsTopologyAndAliasConflicts(t *testing.T) {
	base, snapshot := fixture(t, "p6_nested_form.bin")
	t.Run("rename", func(t *testing.T) {
		base, snapshot := topologyBase(t)
		desired := templateInput(t, base, snapshot)
		oldName := desired.Controls[0].Name
		desired.Controls[0].Name = "Renamed"
		result, err := CompileTemplate(base, desired, 932)
		if err != nil {
			t.Fatal(err)
		}
		projected, err := projection.Project(result)
		if err != nil {
			t.Fatal(err)
		}
		if len(projected.Controls) != len(snapshot.Controls) || findControl(result.Controls, oldName) != nil || findControl(result.Controls, "Renamed") == nil {
			t.Fatalf("rename was not compiled as removal plus addition: %#v", projected.Controls)
		}
	})
	t.Run("reorder", func(t *testing.T) {
		base, snapshot := topologyBase(t)
		desired := templateInput(t, base, snapshot)
		if len(desired.Controls) < 3 || desired.Controls[0].ParentID != "" || desired.Controls[2].ParentID != "" {
			t.Fatal("flat fixture does not have two root siblings")
		}
		desired.Controls[0].ZIndex, desired.Controls[2].ZIndex = new(1), new(0)
		result, err := CompileTemplate(base, desired, 932)
		if err != nil {
			t.Fatal(err)
		}
		if result.Controls[0].Name != snapshot.Controls[2].Name || result.Controls[1].Name != snapshot.Controls[0].Name {
			t.Fatalf("root sibling order = %q, %q", result.Controls[0].Name, result.Controls[1].Name)
		}
	})
	t.Run("duplicate-name", func(t *testing.T) {
		desired := templateInput(t, base, snapshot)
		desired.Controls[1].Name = desired.Controls[0].Name
		_, err := CompileTemplate(base, desired, 932)
		assertTemplateErrorCode(t, err, Invalid)
	})
	t.Run("alias-conflict", func(t *testing.T) {
		desired := templateInput(t, base, snapshot)
		desired.Controls[0].Caption = new("field value")
		desired.Controls[0].Properties = map[string]any{"Caption": "bag value"}
		_, err := CompileTemplate(base, desired, 932)
		assertTemplateErrorCode(t, err, Conflict)
	})
	t.Run("bag-alias-only", func(t *testing.T) {
		desired := templateInput(t, base, snapshot)
		desired.Controls[0].Properties = map[string]any{"caption": "bag only"}
		result, err := CompileTemplate(base, desired, 932)
		if err != nil {
			t.Fatal(err)
		}
		projected, err := projection.Project(result)
		if err != nil {
			t.Fatal(err)
		}
		if projected.Controls[0].Caption == nil || *projected.Controls[0].Caption != "bag only" {
			t.Fatalf("bag-only caption = %#v", projected.Controls[0].Caption)
		}
	})
}

func assertTemplateErrorCode(t *testing.T, err error, code string) {
	t.Helper()
	detail, ok := errors.AsType[*Error](err)
	if !ok || detail.Code != code {
		t.Fatalf("error = %v, want %s", err, code)
	}
}

func TestCompileTemplatePreservesUneditedControlRawBytes(t *testing.T) {
	base, snapshot := fixture(t, "p4_form.bin")
	desired := templateInput(t, base, snapshot)
	desired.Controls[0].Caption = new("changed")
	result, err := CompileTemplate(base, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, original := range base.Controls[1:] {
		updated := findModelControl(result.Controls, original.Name)
		if updated == nil {
			t.Fatalf("control %q disappeared", original.Name)
		}
		if !bytes.Equal(original.Site.Raw, updated.Site.Raw) || !bytes.Equal(original.OpaqueRaw, updated.OpaqueRaw) || !reflect.DeepEqual(original.Record, updated.Record) {
			t.Fatalf("unedited control %q changed", original.Name)
		}
	}
}
