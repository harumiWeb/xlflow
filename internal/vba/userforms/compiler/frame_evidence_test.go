package compiler

import (
	"bytes"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestExcelFramePersistenceEvidence(t *testing.T) {
	read := func(dir, name string) *oforms.Form {
		t.Helper()
		data, err := os.ReadFile(filepath.Join("testdata", dir, name))
		if err != nil {
			t.Fatal(err)
		}
		project, err := vbaproject.Read(data)
		if err != nil {
			t.Fatal(err)
		}
		if len(project.Forms) != 1 {
			t.Fatal("expected one FrameTopologyForm")
		}
		return project.Forms[0]
	}
	baseline := read("frame-excel-authored", "baseline.bin")
	normalized := read("frame-excel-generated", "cli-blank.bin")
	before, err := projection.Project(baseline)
	if err != nil {
		t.Fatal(err)
	}
	generated, err := CompileNew(before, 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, form := range []*oforms.Form{baseline, generated, normalized} {
		got, err := projection.Project(form)
		if err != nil {
			t.Fatal(err)
		}
		if len(got.Controls) != 7 {
			t.Fatal("lost saved/reopened hierarchy")
		}
		for i, want := range before.Controls {
			actual := got.Controls[i]
			if actual.Name != want.Name || actual.Type != want.Type || actual.ParentID != want.ParentID || !reflect.DeepEqual(actual.ZIndex, want.ZIndex) {
				t.Fatalf("saved/reopened hierarchy differs for %s", want.Name)
			}
		}
		for _, name := range []string{"EmptyFrame", "ParentFrame", "NestedFrame"} {
			control := findControl(form.Controls, name)
			excel := findControl(baseline.Controls, name)
			if control.Level.StorageMeta.CLSID != excel.Level.StorageMeta.CLSID || !bytes.Equal(control.Level.CompObjRaw, excel.Level.CompObjRaw) || len(control.Level.ClassTableRaw) != 0 {
				t.Fatalf("Frame identity or omitted class table differs for %s", name)
			}
		}
	}
	for name, cookie := range map[string]int64{"EmptyFrame": 0, "ParentFrame": 3, "NestedFrame": 1} {
		if findControl(generated.Controls, name).Record.Values["ShapeCookie"] != cookie || findControl(baseline.Controls, name).Record.Values["ShapeCookie"] != cookie {
			t.Fatalf("generated Frame descendant count differs from Excel: %s", name)
		}
	}
}

func TestExcelFrameMutationCounterTypeInfoEvidence(t *testing.T) {
	type counters struct {
		shapeCookie     int64
		nextAvailableID int64
	}
	read := func(name string) *oforms.Form {
		t.Helper()
		data, err := os.ReadFile(filepath.Join("testdata", "frame-excel-authored", name+".bin"))
		if err != nil {
			t.Fatal(err)
		}
		project, err := vbaproject.Read(data)
		if err != nil {
			t.Fatal(err)
		}
		if len(project.Forms) != 1 {
			t.Fatalf("%s: expected one FrameTopologyForm, got %d", name, len(project.Forms))
		}
		return project.Forms[0]
	}
	levelCounters := func(form *oforms.Form, owner string) counters {
		t.Helper()
		var level *oforms.Level
		if owner == "" {
			level = form.Levels[0]
		} else {
			control := findControl(form.Controls, owner)
			if control == nil || control.Level == nil {
				t.Fatalf("%s: missing Frame level", owner)
			}
			level = control.Level
		}
		return counters{
			shapeCookie:     level.Record.Values["ShapeCookie"],
			nextAvailableID: level.Record.Values["NextAvailableID"],
		}
	}
	forms := map[string]*oforms.Form{}
	for _, name := range []string{"baseline", "added", "removed", "reordered"} {
		forms[name] = read(name)
		got, found, err := rootFrameTypeInfoVer(forms[name].DesignerSource.Text)
		if err != nil {
			t.Fatalf("%s: read root TypeInfoVer: %v", name, err)
		}
		rootCookie := levelCounters(forms[name], "").shapeCookie
		t.Logf("Excel-authored %s: root TypeInfoVer=%d present=%t ShapeCookie=%d", name, got, found, rootCookie)
		if !found {
			t.Errorf("%s: root VBFrame TypeInfoVer is missing", name)
		} else if got != rootCookie {
			t.Errorf("%s: root TypeInfoVer=%d does not match root ShapeCookie=%d", name, got, rootCookie)
		}
		for _, owner := range []string{"", "ParentFrame", "NestedFrame", "EmptyFrame"} {
			level := levelCounters(forms[name], owner)
			t.Logf("Excel-authored %s owner %q: ShapeCookie=%d NextAvailableID=%d", name, owner, level.shapeCookie, level.nextAvailableID)
		}
	}

	for name, expected := range map[string]map[string]counters{
		"baseline": {
			"":            {shapeCookie: 7, nextAvailableID: 7},
			"ParentFrame": {shapeCookie: 3, nextAvailableID: 7},
			"NestedFrame": {shapeCookie: 1, nextAvailableID: 7},
			"EmptyFrame":  {shapeCookie: 0, nextAvailableID: 0},
		},
		"added": {
			"":            {shapeCookie: 8, nextAvailableID: 8},
			"ParentFrame": {shapeCookie: 4, nextAvailableID: 8},
			"NestedFrame": {shapeCookie: 2, nextAvailableID: 8},
		},
		"removed": {
			"":            {shapeCookie: 9, nextAvailableID: 8},
			"ParentFrame": {shapeCookie: 5, nextAvailableID: 8},
			"NestedFrame": {shapeCookie: 3, nextAvailableID: 8},
		},
		"reordered": {
			"":            {shapeCookie: 9, nextAvailableID: 8},
			"ParentFrame": {shapeCookie: 5, nextAvailableID: 8},
			"NestedFrame": {shapeCookie: 3, nextAvailableID: 8},
			"EmptyFrame":  {shapeCookie: 0, nextAvailableID: 0},
		},
	} {
		for owner, want := range expected {
			if got := levelCounters(forms[name], owner); got != want {
				t.Errorf("Excel-authored %s owner %q counters = %+v, want %+v", name, owner, got, want)
			}
		}
	}
	diffLabel := findControl(forms["added"].Controls, "DiffLabel")
	nestedFrame := findControl(forms["added"].Controls, "NestedFrame")
	diffLabelUnderNested := false
	if nestedFrame != nil {
		for _, child := range nestedFrame.Children {
			diffLabelUnderNested = diffLabelUnderNested || child == diffLabel
		}
	}
	if diffLabel == nil || diffLabel.ID != 8 || !diffLabelUnderNested {
		t.Errorf("Excel-authored added DiffLabel should be ID 8 under NestedFrame; got control=%#v NestedFrame=%#v", diffLabel, nestedFrame)
	}
	if got := findControl(forms["removed"].Controls, "DiffLabel"); got != nil {
		t.Errorf("Excel-authored removed fixture still contains DiffLabel ID %d", got.ID)
	}
	parent := findControl(forms["reordered"].Controls, "ParentFrame")
	if parent == nil || len(parent.Children) != 2 || parent.Children[0].Name != "NestedFrame" || parent.Children[1].Name != "SiblingText" {
		t.Errorf("Excel-authored reordered ParentFrame children are not [NestedFrame SiblingText]")
	}
	for _, owner := range []string{"", "ParentFrame", "NestedFrame", "EmptyFrame"} {
		if got, want := levelCounters(forms["reordered"], owner), levelCounters(forms["removed"], owner); got.shapeCookie != want.shapeCookie {
			t.Errorf("Excel-authored sibling reorder changed %q ShapeCookie: removed=%d reordered=%d", owner, want.shapeCookie, got.shapeCookie)
		}
	}

	baselineSpec, err := projection.Project(forms["baseline"])
	if err != nil {
		t.Fatal(err)
	}
	addedSpec, err := projection.Project(forms["added"])
	if err != nil {
		t.Fatal(err)
	}
	removedSpec, err := projection.Project(forms["removed"])
	if err != nil {
		t.Fatal(err)
	}
	reorderedSpec, err := projection.Project(forms["reordered"])
	if err != nil {
		t.Fatal(err)
	}
	var siblingZIndex *int
	for _, control := range reorderedSpec.Controls {
		if control.Name == "SiblingText" {
			siblingZIndex = control.ZIndex
			break
		}
	}
	if siblingZIndex == nil || *siblingZIndex != 1 {
		t.Errorf("Excel-authored reordered SiblingText ZIndex=%v, want 1", siblingZIndex)
	}
	compile := func(base *oforms.Form, before, after spec.FormSpec) *oforms.Form {
		t.Helper()
		compiled, err := CompileEdits(base, before, after, 932)
		if err != nil {
			t.Fatal(err)
		}
		return compiled
	}
	compareResult := func(name string, actual, excel *oforms.Form, want spec.FormSpec) {
		t.Helper()
		got, err := projection.Project(actual)
		if err != nil {
			t.Fatal(err)
		}
		if len(got.Controls) != len(want.Controls) {
			t.Errorf("compiler %s control count=%d, Excel fixture=%d", name, len(got.Controls), len(want.Controls))
		} else {
			for i, expected := range want.Controls {
				actual := got.Controls[i]
				if actual.Name != expected.Name || actual.Type != expected.Type || actual.ParentID != expected.ParentID || !reflect.DeepEqual(actual.ZIndex, expected.ZIndex) {
					t.Errorf("compiler %s topology differs at %s", name, expected.Name)
				}
			}
		}
		compiledTypeInfo, found, err := rootFrameTypeInfoVer(actual.DesignerSource.Text)
		if err != nil {
			t.Errorf("compiler %s: read root TypeInfoVer: %v", name, err)
		} else if !found || compiledTypeInfo != levelCounters(actual, "").shapeCookie {
			t.Errorf("compiler %s root TypeInfoVer=%d present=%t, ShapeCookie=%d", name, compiledTypeInfo, found, levelCounters(actual, "").shapeCookie)
		}
		if got, want := levelCounters(actual, "").shapeCookie, levelCounters(excel, "").shapeCookie; got != want {
			t.Errorf("compiler %s root ShapeCookie=%d, Excel fixture=%d", name, got, want)
		}
		t.Logf("compiler %s: root TypeInfoVer=%d ShapeCookie=%d NextAvailableID=%d; Excel root ShapeCookie=%d NextAvailableID=%d", name, compiledTypeInfo, levelCounters(actual, "").shapeCookie, levelCounters(actual, "").nextAvailableID, levelCounters(excel, "").shapeCookie, levelCounters(excel, "").nextAvailableID)
		for _, owner := range []string{"ParentFrame", "NestedFrame", "EmptyFrame"} {
			gotCounters, wantCounters := levelCounters(actual, owner), levelCounters(excel, owner)
			t.Logf("compiler %s %s: ShapeCookie=%d NextAvailableID=%d; Excel ShapeCookie=%d NextAvailableID=%d", name, owner, gotCounters.shapeCookie, gotCounters.nextAvailableID, wantCounters.shapeCookie, wantCounters.nextAvailableID)
			if gotCounters.shapeCookie != wantCounters.shapeCookie {
				t.Errorf("compiler %s %s ShapeCookie=%d, Excel fixture=%d", name, owner, gotCounters.shapeCookie, wantCounters.shapeCookie)
			}
		}
	}

	compiledAdded := compile(forms["baseline"], baselineSpec, addedSpec)
	compareResult("add", compiledAdded, forms["added"], addedSpec)
	// Excel's added.bin stores NextAvailableID=8 for DiffLabel ID 8. The
	// compiler keeps its existing one-past high-water contract, so its next ID
	// is 9; the fixture's counter is logged but is not an equality assertion.
	newControl := findControl(compiledAdded.Controls, "DiffLabel")
	maxID := int64(0)
	for _, level := range compiledAdded.Levels {
		for _, control := range level.Controls {
			maxID = max(maxID, int64(control.ID))
		}
	}
	if newControl == nil || newControl.ID != 8 || levelCounters(compiledAdded, "").nextAvailableID != maxID+1 {
		t.Errorf("compiler add must assign DiffLabel ID 8 and retain one-past NextAvailableID: control=%v next=%d maxID=%d", newControl, levelCounters(compiledAdded, "").nextAvailableID, maxID)
	}

	compiledAddedSpec, err := projection.Project(compiledAdded)
	if err != nil {
		t.Fatal(err)
	}
	compiledRemoved := compile(compiledAdded, compiledAddedSpec, removedSpec)
	compareResult("remove", compiledRemoved, forms["removed"], removedSpec)
	if levelCounters(compiledRemoved, "").nextAvailableID < levelCounters(compiledAdded, "").nextAvailableID {
		t.Errorf("compiler removal regressed root NextAvailableID from %d to %d", levelCounters(compiledAdded, "").nextAvailableID, levelCounters(compiledRemoved, "").nextAvailableID)
	}

	compiledRemovedSpec, err := projection.Project(compiledRemoved)
	if err != nil {
		t.Fatal(err)
	}
	compiledReordered := compile(compiledRemoved, compiledRemovedSpec, reorderedSpec)
	compareResult("reorder", compiledReordered, forms["reordered"], reorderedSpec)
	compiledParent := findControl(compiledReordered.Controls, "ParentFrame")
	if compiledParent == nil || len(compiledParent.Children) != 2 || compiledParent.Children[0].Name != "NestedFrame" || compiledParent.Children[1].Name != "SiblingText" {
		t.Errorf("compiler reorder did not produce ParentFrame children [NestedFrame SiblingText]")
	}
}

func rootFrameTypeInfoVer(source string) (int64, bool, error) {
	depth := 0
	var value int64
	found := false
	for _, line := range strings.Split(source, "\n") {
		trimmed := strings.TrimSpace(strings.TrimSuffix(line, "\r"))
		if depth == 1 && strings.HasPrefix(strings.ToLower(trimmed), "typeinfover") {
			if found {
				return 0, false, strconv.ErrSyntax
			}
			_, raw, ok := strings.Cut(trimmed, "=")
			if !ok {
				return 0, false, strconv.ErrSyntax
			}
			raw, _, _ = strings.Cut(raw, "'")
			parsed, err := strconv.ParseInt(strings.TrimSpace(raw), 10, 64)
			if err != nil {
				return 0, false, err
			}
			value, found = parsed, true
		}
		lower := strings.ToLower(trimmed)
		if strings.HasPrefix(lower, "begin ") {
			depth++
		} else if lower == "end" {
			depth--
		}
	}
	return value, found, nil
}
