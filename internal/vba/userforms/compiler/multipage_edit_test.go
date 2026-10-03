package compiler

import (
	"errors"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestTabSelectionFollowsRetainedIdentityAndExplicitTemplateIndex(t *testing.T) {
	for _, mode := range []string{"edits", "template implicit", "template explicit"} {
		t.Run(mode, func(t *testing.T) {
			base, before := multipageExcelAuthoredFixture(t, "baseline")
			desired := snapshotCopy(t, before)
			alpha := projectedControlByName(t, desired, "PageAlpha")
			beta := projectedControlByName(t, desired, "PageBeta")
			alpha.ZIndex = new(1)
			beta.ZIndex = new(0)
			strip := projectedControlByName(t, desired, "StripMain")
			strip.Tabs[0], strip.Tabs[1] = strip.Tabs[1], strip.Tabs[0]
			if mode == "template implicit" {
				strip.SelectedIndex = nil
				projectedControlByName(t, desired, "MultiMain").SelectedIndex = nil
			}
			var result *oforms.Form
			var err error
			if mode == "edits" {
				result, err = CompileEdits(base, before, desired, 932)
			} else {
				result, err = CompileTemplate(base, desired, 932)
			}
			if err != nil {
				t.Fatal(err)
			}
			actual, err := projection.Project(result)
			if err != nil {
				t.Fatal(err)
			}
			want := 0
			if mode == "template explicit" {
				want = 1
			}
			for _, name := range []string{"MultiMain", "StripMain"} {
				got := projectedControlByName(t, actual, name)
				if got.SelectedIndex == nil || *got.SelectedIndex != want {
					t.Fatalf("%s selection=%v want%d", name, got.SelectedIndex, want)
				}
			}
			assertMultiPageSiteSelection(t, result, "MultiMain")
		})
	}
}

func assertMultiPageSiteSelection(t *testing.T, form *oforms.Form, name string) {
	t.Helper()
	owner := findModelControl(form.Controls, name)
	for i, page := range owner.MultiPage.Pages {
		visible := page.Site.Values["BitFlags"]&2 != 0
		if want := int32(i) == owner.MultiPage.Hidden.TabStrip.SelectedIndex; visible != want {
			t.Fatalf("%s Site visible=%t want=%t (selected=%d)", page.Name, visible, want, owner.MultiPage.Hidden.TabStrip.SelectedIndex)
		}
	}
}

func TestMultiPageTopologySynchronizesPageSiteSelection(t *testing.T) {
	for _, mode := range []string{"remove selected", "add unselected", "remove all"} {
		t.Run(mode, func(t *testing.T) {
			input := spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: "SelectionForm"}, Controls: []spec.FormSpecControl{
				{ID: "multi", Name: "Pages", Type: "MultiPage", SelectedIndex: new(1), ZIndex: new(0)},
				{ID: "a", ParentID: "multi", Name: "PageA", Type: "Page", ZIndex: new(0)},
				{ID: "b", ParentID: "multi", Name: "PageB", Type: "Page", ZIndex: new(1)},
			}}
			base, err := CompileNew(input, 932)
			if err != nil {
				t.Fatal(err)
			}
			before, err := projection.Project(base)
			if err != nil {
				t.Fatal(err)
			}
			desired := snapshotCopy(t, before)
			wantIndex := int32(1)
			switch mode {
			case "remove selected":
				desired.Controls = desired.Controls[:2]
				desired.Controls[0].SelectedIndex = new(0)
				desired.Controls[0].Observed.SelectedIndex = new(0)
				wantIndex = 0
			case "add unselected":
				desired.Controls = append(desired.Controls, spec.FormSpecControl{ID: "c", ParentID: desired.Controls[0].ID, Name: "PageC", Type: "Page", ZIndex: new(2)})
			case "remove all":
				desired.Controls = desired.Controls[:1]
				desired.Controls[0].SelectedIndex = nil
				desired.Controls[0].Observed = nil
				wantIndex = -1
			}
			result, err := CompileEdits(base, before, desired, 932)
			if err != nil {
				t.Fatal(err)
			}
			if got := findModelControl(result.Controls, "Pages").MultiPage.Hidden.TabStrip.SelectedIndex; got != wantIndex {
				t.Fatalf("selection=%d want=%d", got, wantIndex)
			}
			assertMultiPageSiteSelection(t, result, "Pages")
			assertMultiPageSiteSelection(t, base, "Pages")
		})
	}
}

func TestRemovingSelectedTabUsesFirstRemainingAndEmptySelection(t *testing.T) {
	for _, empty := range []bool{false, true} {
		base, before := multipageExcelAuthoredFixture(t, "baseline")
		desired := snapshotCopy(t, before)
		strip := projectedControlByName(t, desired, "StripMain")
		if empty {
			strip.Tabs = []spec.FormSpecTab{}
		} else {
			strip.Tabs = strip.Tabs[:1]
		}
		// The old numeric selection is unchanged in this before/after edit; the
		// collection change must reconcile identity before validating the final index.
		result, err := CompileEdits(base, before, desired, 932)
		if err != nil {
			t.Fatal(err)
		}
		actual, err := projection.Project(result)
		if err != nil {
			t.Fatal(err)
		}
		want := 0
		if empty {
			want = -1
		}
		if got := projectedControlByName(t, actual, "StripMain"); got.SelectedIndex == nil || *got.SelectedIndex != want {
			t.Fatalf("selection=%v want%d", got.SelectedIndex, want)
		}
	}
}

func TestTabEditRejectsStaleStateAndInvalidSelectionWithoutMutation(t *testing.T) {
	for _, which := range []string{"stale tabs", "invalid selection", "page geometry"} {
		t.Run(which, func(t *testing.T) {
			base, before := multipageExcelAuthoredFixture(t, "baseline")
			original, err := oforms.SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			desired := snapshotCopy(t, before)
			want := Invalid
			switch which {
			case "stale tabs":
				projectedControlByName(t, before, "StripMain").Tabs[0].Caption = new("stale")
				projectedControlByName(t, desired, "StripMain").Tabs[0].Caption = new("edited")
				want = Stale
			case "invalid selection":
				projectedControlByName(t, desired, "StripMain").SelectedIndex = new(99)
			case "page geometry":
				projectedControlByName(t, desired, "PageAlpha").Width = new(100.0)
			}
			beforeCopy, desiredCopy := snapshotCopy(t, before), snapshotCopy(t, desired)
			result, err := CompileEdits(base, before, desired, 932)
			detail, ok := errors.AsType[*Error](err)
			if result != nil || !ok || detail.Code != want {
				t.Fatalf("result=%v err=%v want%s", result, err, want)
			}
			unchanged, err := oforms.SerializeForm(base, 932)
			if err != nil || !reflect.DeepEqual(original, unchanged) {
				t.Fatal("failure mutated persistence")
			}
			if !reflect.DeepEqual(beforeCopy, before) || !reflect.DeepEqual(desiredCopy, desired) {
				t.Fatal("failure mutated authoring input")
			}
		})
	}
}

func TestEmptyMultiPageRejectsDirectCachedTabMutation(t *testing.T) {
	base, _ := multipageExcelAuthoredFixture(t, "empty")
	owner := findModelControl(base.Controls, "MultiMain")
	if len(owner.MultiPage.Hidden.TabStrip.Tabs) == 0 {
		t.Fatal("fixture must retain ghost tabs")
	}
	original, err := oforms.SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := oforms.ApplyTabEdits(base, []oforms.TabEdit{{Control: owner.Name, State: &oforms.TabStrip{SelectedIndex: -1}}}, 932)
	if got != nil || !errors.Is(err, oforms.ErrUnsupportedEdit) {
		t.Fatalf("got=%v err=%v", got, err)
	}
	unchanged, err := oforms.SerializeForm(base, 932)
	if err != nil || !reflect.DeepEqual(original, unchanged) {
		t.Fatal("rejected ghost clear mutated input")
	}
}
