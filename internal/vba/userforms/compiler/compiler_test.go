package compiler

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func fixture(t *testing.T, name string) (*oforms.Form, spec.FormSpec) {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", name))
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	form, err := oforms.ReadForm(container, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	snapshot, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	return form, snapshot
}

func snapshotCopy(t *testing.T, input spec.FormSpec) spec.FormSpec {
	t.Helper()
	body, err := json.Marshal(input)
	if err != nil {
		t.Fatal(err)
	}
	var result spec.FormSpec
	if err := json.Unmarshal(body, &result); err != nil {
		t.Fatal(err)
	}
	return result
}

func TestCompileNoopAndMetadataPreserveAllBytes(t *testing.T) {
	for _, name := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		t.Run(name, func(t *testing.T) {
			form, before := fixture(t, name)
			after := snapshotCopy(t, before)
			after.Warnings = nil
			after.Form.Observed = &spec.FormSpecObservedForm{Caption: new("not build intent")}
			for i := range after.Controls {
				after.Controls[i].Observed = nil
				after.Controls[i].Unsupported = nil
			}
			result, err := CompileEdits(form, before, after, 932)
			if err != nil {
				t.Fatal(err)
			}
			a, err := oforms.SerializeForm(form, 932)
			if err != nil {
				t.Fatal(err)
			}
			b, err := oforms.SerializeForm(result, 932)
			if err != nil {
				t.Fatal(err)
			}
			assertDesignerBytes(t, a, b)
		})
	}
}

func TestCompileCaptionGrowthGeometryAndProjectRoundTrip(t *testing.T) {
	form, before := fixture(t, "p4_form.bin")
	original, err := oforms.SerializeForm(form, 932)
	if err != nil {
		t.Fatal(err)
	}
	after := snapshotCopy(t, before)
	after.Form.Caption = new("編集フォーム")
	after.Controls[0].Caption = new(strings.Repeat("日本語 caption ", 20))
	after.Controls[0].Left = new(12.25)
	after.Controls[0].Top = new(-3.5)
	after.Controls[0].Width = new(123.5)
	after.Controls[0].Height = new(42.25)
	after.Controls[0].Visible = new(false)
	after.Controls[0].TabIndex = new(0)
	first, err := CompileEdits(form, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	second, err := CompileEdits(form, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	serialized, err := oforms.SerializeForm(first, 932)
	if err != nil {
		t.Fatal(err)
	}
	repeated, err := oforms.SerializeForm(second, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(serialized, repeated) {
		t.Fatal("non-deterministic compilation")
	}
	unchanged, err := oforms.SerializeForm(form, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(original, unchanged) {
		t.Fatal("compiler changed input model")
	}
	for _, control := range form.Controls[1:] {
		var matched *oforms.Control
		for _, updated := range first.Controls {
			if control.Name == updated.Name {
				matched = updated
			}
		}
		if matched == nil {
			t.Fatal("lost unedited control")
			return
		}
		if !bytes.Equal(control.Site.Raw, matched.Site.Raw) || !reflect.DeepEqual(control.Record, matched.Record) || !bytes.Equal(control.OpaqueRaw, matched.OpaqueRaw) {
			t.Fatalf("unedited control %s changed", control.Name)
		}
	}
	projectBody, err := os.ReadFile(filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", "p4_form.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(projectBody)
	if err != nil {
		t.Fatal(err)
	}
	project.Forms[0] = first
	output, err := vbaproject.Write(project)
	if err != nil {
		t.Fatal(err)
	}
	readBack, err := vbaproject.Read(output)
	if err != nil {
		t.Fatal(err)
	}
	got, err := projection.Project(readBack.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	if got.Form.Caption == nil || *got.Form.Caption != *after.Form.Caption || got.Controls[0].Caption == nil || *got.Controls[0].Caption != *after.Controls[0].Caption {
		t.Fatal("caption did not survive project writer")
	}
	left, _ := convertValue("CommandButton", "left", *got.Controls[0].Left)
	wantLeft, _ := convertValue("CommandButton", "left", *after.Controls[0].Left)
	if left != wantLeft || *got.Controls[0].Visible {
		t.Fatal("geometry/visibility did not survive")
	}
	// Shrinking to an explicit empty string must retain a valid site extent.
	shrunk := snapshotCopy(t, got)
	shrunk.Controls[0].Caption = new("")
	if _, err := CompileEdits(readBack.Forms[0], got, shrunk, 932); err != nil {
		t.Fatal(err)
	}
}

func TestCompileRejectsEditsAtomically(t *testing.T) {
	for _, tc := range []struct {
		name, code string
		change     func(*spec.FormSpec, *spec.FormSpec)
	}{
		{"root-size", Invalid, func(_, a *spec.FormSpec) { a.Form.Width = new(300.0) }},
		{"rename", Unsupported, func(_, a *spec.FormSpec) { a.Controls[0].Name = "Renamed" }},
		{"remove", Unsupported, func(_, a *spec.FormSpec) { a.Controls = a.Controls[1:] }},
		{"stale", Stale, func(b, a *spec.FormSpec) {
			b.Controls[0].Caption = new("wrong baseline")
			a.Controls[0].Caption = new("requested")
		}},
		{"property-reset", Unsupported, func(_, a *spec.FormSpec) { a.Controls[0].Caption = nil }},
		{"unknown-property", Unsupported, func(_, a *spec.FormSpec) { a.Controls[0].Properties = map[string]any{"UnmappedProperty": 1} }},
		{"invalid-property", Invalid, func(_, a *spec.FormSpec) { a.Controls[0].Properties = map[string]any{"BackColor": -1} }},
		{"alias-conflict", Conflict, func(_, a *spec.FormSpec) {
			a.Controls[0].Caption = new("one")
			a.Controls[0].Properties = map[string]any{"Caption": "two"}
		}},
		{"tab-range", Invalid, func(_, a *spec.FormSpec) { a.Controls[0].TabIndex = new(32768) }},
		{"geometry-range", Invalid, func(_, a *spec.FormSpec) { a.Controls[0].Left = new(1e20) }},
		{"before-coordinate-system", Invalid, func(b, a *spec.FormSpec) {
			b.CoordinateSystem = "pixels"
			a.Controls[0].Left = new(72.0)
		}},
		{"after-coordinate-system", Invalid, func(_, a *spec.FormSpec) {
			a.CoordinateSystem = "pixels"
			a.Controls[0].Properties = map[string]any{"Left": 72.0}
		}},
		{"new-build-caption-stale-legacy", Stale, func(b, a *spec.FormSpec) {
			b.Form.Build = nil
			b.Form.Caption = new("stale legacy caption")
			a.Form.Build = &spec.FormSpecBuildForm{Caption: new("requested")}
		}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			form, before := fixture(t, "p4_form.bin")
			original, err := oforms.SerializeForm(form, 932)
			if err != nil {
				t.Fatal(err)
			}
			after := snapshotCopy(t, before)
			tc.change(&before, &after)
			beforeBytes, _ := json.Marshal(before)
			afterBytes, _ := json.Marshal(after)
			result, err := CompileEdits(form, before, after, 932)
			detail, ok := errors.AsType[*Error](err)
			if result != nil || !ok || detail.Code != tc.code || detail.Property == "" {
				t.Fatalf("got result=%v err=%v, want %s", result != nil, err, tc.code)
			}
			current, err := oforms.SerializeForm(form, 932)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(current, original) {
				t.Fatal("failure mutated base")
			}
			b, _ := json.Marshal(before)
			a, _ := json.Marshal(after)
			if !bytes.Equal(beforeBytes, b) || !bytes.Equal(afterBytes, a) {
				t.Fatal("failure mutated snapshots")
			}
		})
	}
}

func TestCompileExplicitBuildCaptionAndBagAlias(t *testing.T) {
	form, before := fixture(t, "p4_form.bin")
	after := snapshotCopy(t, before)
	if after.Form.Build == nil {
		after.Form.Build = &spec.FormSpecBuildForm{}
	}
	after.Form.Build.Caption = new("build caption")
	after.Form.Caption = new("legacy caption")
	after.Controls[0].Caption = new("same")
	after.Controls[0].Properties = map[string]any{"Caption": "same", "Tag": "日本語", "ControlTipText": "tip", "BackColor": float64(0x8000000f)}
	result, err := CompileEdits(form, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got := result.Levels[0].Record.Strings["Caption"].Text; got != "build caption" {
		t.Fatalf("caption=%q", got)
	}
	if got := result.Controls[0].Site.Strings["Tag"].Text; got != "日本語" {
		t.Fatalf("tag=%q", got)
	}
}

func TestCompileCoordinateSystems(t *testing.T) {
	for _, beforeBasis := range []string{"", "points", "parent-relative"} {
		for _, afterBasis := range []string{"", "points", "parent-relative"} {
			t.Run(beforeBasis+"/"+afterBasis, func(t *testing.T) {
				form, before := fixture(t, "p6_nested_form.bin")
				before.CoordinateSystem = beforeBasis
				after := snapshotCopy(t, before)
				after.CoordinateSystem = afterBasis
				index := -1
				for i, control := range after.Controls {
					if control.Type == "TextBox" && control.ParentID != "" {
						index = i
						break
					}
				}
				if index < 0 {
					t.Fatal("fixture missing nested TextBox")
				}
				after.Controls[index].Left = new(72.0)
				after.Controls[index].Properties = map[string]any{"Top": 36.0}
				result, err := CompileEdits(form, before, after, 932)
				if err != nil {
					t.Fatal(err)
				}
				control := findControl(result.Controls, after.Controls[index].Name)
				if control == nil || control.Site.Position == nil || control.Site.Position.Left != 2540 || control.Site.Position.Top != 1270 {
					t.Fatal("coordinates must remain parent-relative points")
				}
			})
		}
	}
	for _, input := range []string{"before", "after"} {
		t.Run(input+"-invalid-noop", func(t *testing.T) {
			form, before := fixture(t, "p4_form.bin")
			after := snapshotCopy(t, before)
			if input == "before" {
				before.CoordinateSystem = "pixels"
			} else {
				after.CoordinateSystem = "pixels"
			}
			result, err := CompileEdits(form, before, after, 932)
			detail, ok := errors.AsType[*Error](err)
			if result != nil || !ok || detail.Code != Invalid || detail.Property != input+".coordinateSystem" {
				t.Fatalf("expected invalid coordinate system: %v", err)
			}
		})
	}
}

func TestCompileNewBuildCaptionChecksLegacyBaseline(t *testing.T) {
	for _, tc := range []struct {
		name, wantCode string
		change         func(*spec.FormSpec, *spec.FormSpec)
	}{
		{"matching-legacy", "", func(_, _ *spec.FormSpec) {}},
		{"empty-build-matching-legacy", "", func(b, _ *spec.FormSpec) { b.Form.Build = &spec.FormSpecBuildForm{} }},
		{"stale-legacy", Stale, func(b, _ *spec.FormSpec) { b.Form.Caption = new("stale") }},
		{"empty-build-stale-legacy", Stale, func(b, _ *spec.FormSpec) {
			b.Form.Build = &spec.FormSpecBuildForm{}
			b.Form.Caption = new("stale")
		}},
		{"stale-legacy-equal-new-build", Stale, func(b, a *spec.FormSpec) { b.Form.Caption = a.Form.Build.Caption }},
		{"absent-legacy", "", func(b, _ *spec.FormSpec) { b.Form.Caption = nil }},
		{"existing-build-precedence", "", func(b, _ *spec.FormSpec) {
			b.Form.Build = &spec.FormSpecBuildForm{Caption: b.Form.Caption}
			b.Form.Caption = new("ignored legacy baseline")
		}},
		{"stale-existing-build", Stale, func(b, _ *spec.FormSpec) {
			b.Form.Build = &spec.FormSpecBuildForm{Caption: new("stale build")}
		}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			form, before := fixture(t, "p4_form.bin")
			before.Form.Build = nil
			if before.Form.Caption == nil {
				before.Form.Caption = new("")
			}
			after := snapshotCopy(t, before)
			after.Form.Build = &spec.FormSpecBuildForm{Caption: new("requested build caption")}
			after.Form.Caption = new("ignored legacy edit")
			tc.change(&before, &after)
			result, err := CompileEdits(form, before, after, 932)
			if tc.wantCode != "" {
				detail, ok := errors.AsType[*Error](err)
				if result != nil || !ok || detail.Code != tc.wantCode || detail.Property != "form.caption" {
					t.Fatalf("expected stale caption: %v", err)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if got := result.Levels[0].Record.Strings["Caption"].Text; got != *after.Form.Build.Caption {
				t.Fatalf("build precedence lost: %q", got)
			}
		})
	}
}

func TestCompilePropertyKeyUnionIsDeterministic(t *testing.T) {
	for _, tc := range []struct {
		name           string
		previous, next map[string]any
	}{
		{"empty", nil, nil},
		{"overlap", map[string]any{"Tag": "", "ControlTipText": ""}, map[string]any{"Tag": "tag", "ControlTipText": "tip", "Caption": "caption"}},
		{"disjoint", map[string]any{"Tag": nil}, map[string]any{"ControlTipText": "tip", "Caption": "caption"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			form, before := fixture(t, "p4_form.bin")
			before.Controls[0].Properties = tc.previous
			after := snapshotCopy(t, before)
			after.Controls[0].Properties = tc.next
			var original *oforms.SerializedForm
			for range 5 {
				result, err := CompileEdits(form, before, after, 932)
				if err != nil {
					t.Fatal(err)
				}
				serialized, err := oforms.SerializeForm(result, 932)
				if err != nil {
					t.Fatal(err)
				}
				for property, want := range tc.next {
					got, found := editedProperty(result.Controls[0], property)
					if !found || !reflect.DeepEqual(got, want) {
						t.Fatalf("property %s: got %v, want %v", property, got, want)
					}
				}
				if original != nil {
					assertDesignerBytes(t, original, serialized)
				}
				original = serialized
			}
		})
	}
}

func TestCompileNestedTextAliasesAndUneditedSubtrees(t *testing.T) {
	form, before := fixture(t, "p6_nested_form.bin")
	index := -1
	originalControls := flattenControls(form.Controls)
	for i, control := range before.Controls {
		if control.Type == "TextBox" && control.ParentID != "" {
			index = i
			break
		}
	}
	if index == -1 {
		t.Fatal("fixture missing nested TextBox")
	}
	after := snapshotCopy(t, before)
	after.Controls[index].Text = new("日本語 text 🐚")
	after.Controls[index].Left = new(7.25)
	after.Controls[index].Properties = map[string]any{"Tag": "nested"}
	compiled, err := CompileEdits(form, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	updatedControls := flattenControls(compiled.Controls)
	got, err := projection.Project(compiled)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[index].Text == nil || *got.Controls[index].Text != *after.Controls[index].Text || got.Controls[index].Value != *after.Controls[index].Text {
		t.Fatal("Text must update persisted Value, retaining no stale alias")
	}
	for i, control := range before.Controls {
		if i == index {
			continue
		}
		original := originalControls[i]
		updated := updatedControls[i]
		if original == nil || updated == nil {
			t.Fatal("missing control")
			return
		}
		if !bytes.Equal(original.Site.Raw, updated.Site.Raw) || !reflect.DeepEqual(original.Record, updated.Record) {
			t.Fatalf("unedited control %s changed", control.Name)
		}
	}
	conflicting := snapshotCopy(t, before)
	conflicting.Controls[index].Text = new("one")
	conflicting.Controls[index].Value = "two"
	_, err = CompileEdits(form, before, conflicting, 932)
	detail, ok := errors.AsType[*Error](err)
	if !ok || detail.Code != Conflict {
		t.Fatalf("expected alias conflict, got %v", err)
	}
}

func assertDesignerBytes(t *testing.T, a, b *oforms.SerializedForm) {
	t.Helper()
	if !reflect.DeepEqual(a.Storages, b.Storages) || len(a.Streams) != len(b.Streams) {
		t.Fatal("Designer storage metadata or stream inventory changed")
	}
	for path, raw := range a.Streams {
		actual, exists := b.Streams[path]
		if !exists || !bytes.Equal(raw, actual) {
			t.Fatalf("Designer stream %q changed", path)
		}
	}
}

func flattenControls(controls []*oforms.Control) []*oforms.Control {
	var result []*oforms.Control
	for _, control := range controls {
		result = append(result, control)
		result = append(result, flattenControls(control.Children)...)
	}
	return result
}

func findControl(controls []*oforms.Control, name string) *oforms.Control {
	for _, control := range controls {
		if control.Name == name {
			return control
		}
		if child := findControl(control.Children, name); child != nil {
			return child
		}
	}
	return nil
}

func excelProject(t *testing.T, name string) *vbaproject.Project {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("testdata", "excel-authored", name))
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(body)
	if err != nil {
		t.Fatal(err)
	}
	return project
}

func TestExcelAuthoredEnabledPersistenceEvidence(t *testing.T) {
	baseline := excelProject(t, "00-baseline.bin")
	disabled := excelProject(t, "04-all-disabled.bin")
	for _, control := range flattenControls(baseline.Forms[0].Controls) {
		other := findControl(disabled.Forms[0].Controls, control.Name)
		if other == nil {
			t.Fatal("disabled fixture lost control")
			return
		}
		before, err := projection.Project(baseline.Forms[0])
		if err != nil {
			t.Fatal(err)
		}
		after := snapshotCopy(t, before)
		for i := range after.Controls {
			if after.Controls[i].Name == control.Name {
				after.Controls[i].Enabled = new(false)
			}
		}
		compiled, err := CompileEdits(baseline.Forms[0], before, after, baseline.Props.CodePage)
		if err != nil {
			t.Fatal(err)
		}
		changed := findControl(compiled.Controls, control.Name)
		field := "VariousPropertyBits"
		if control.Kind == "MSForms.Frame" {
			field = "BooleanProperties"
		}
		if changed.Record.Values[field] != other.Record.Values[field] {
			t.Fatalf("%s %s differs from Excel persisted value", control.Name, field)
		}
		got, err := projection.Project(compiled)
		if err != nil {
			t.Fatal(err)
		}
		for _, item := range got.Controls {
			if item.Name == control.Name && (item.Enabled == nil || *item.Enabled) {
				t.Fatalf("disabled %s projected as enabled", control.Name)
			}
		}
	}
}

func TestCompileAcceptsExcelTwipRoundingButRejectsStaleGeometry(t *testing.T) {
	form, before := fixture(t, "p4_form.bin")
	before.Controls[0].Left = new(*before.Controls[0].Left + 0.02)
	after := snapshotCopy(t, before)
	after.Controls[0].Left = new(*before.Controls[0].Left + 3)
	if _, err := CompileEdits(form, before, after, 932); err != nil {
		t.Fatal(err)
	}
	before.Controls[0].Left = new(*before.Controls[0].Left + 1)
	_, err := CompileEdits(form, before, after, 932)
	detail, ok := errors.AsType[*Error](err)
	if !ok || detail.Code != Stale {
		t.Fatalf("expected stale geometry error: %v", err)
	}
}

func TestCompileRejectsNonPersistedListState(t *testing.T) {
	project := excelProject(t, "00-baseline.bin")
	before, err := projection.Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	for _, property := range []string{"list", "selectedIndex", "value"} {
		t.Run(property, func(t *testing.T) {
			after := snapshotCopy(t, before)
			for i := range after.Controls {
				if after.Controls[i].Type != "ListBox" {
					continue
				}
				switch property {
				case "list":
					after.Controls[i].List = []string{"alpha"}
				case "selectedIndex":
					after.Controls[i].SelectedIndex = new(1)
				case "value":
					after.Controls[i].Value = "alpha"
				}
			}
			_, err := CompileEdits(project.Forms[0], before, after, project.Props.CodePage)
			detail, ok := errors.AsType[*Error](err)
			if !ok || detail.Code != Unsupported {
				t.Fatalf("expected unsupported list state: %v", err)
			}
		})
	}
}

func TestCompileRejectsInapplicableDirectCaption(t *testing.T) {
	project := excelProject(t, "00-baseline.bin")
	before, err := projection.Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	for _, kind := range []string{"TextBox", "ComboBox", "ListBox"} {
		t.Run(kind, func(t *testing.T) {
			after := snapshotCopy(t, before)
			for i := range after.Controls {
				if after.Controls[i].Type == kind {
					after.Controls[i].Caption = new("inapplicable")
					break
				}
			}
			result, err := CompileEdits(project.Forms[0], before, after, project.Props.CodePage)
			detail, ok := errors.AsType[*Error](err)
			if result != nil || !ok || detail.Code != Unsupported || !strings.HasSuffix(detail.Property, ".caption") {
				t.Fatalf("want unsupported caption, got %v", err)
			}
		})
	}
}
