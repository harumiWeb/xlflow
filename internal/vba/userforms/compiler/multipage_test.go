package compiler

import (
	"bytes"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestCompileTemplateExcelAuthoredMultiPageNoOpAndEmptySelection(t *testing.T) {
	for _, fixtureName := range []string{"baseline", "empty"} {
		t.Run(fixtureName, func(t *testing.T) {
			base, before := multipageExcelAuthoredFixture(t, fixtureName)
			if fixtureName == "baseline" {
				assertAuthoredPageMetadata(t, before)
				assertAuthoredTabMetadata(t, before)
			} else {
				multiPage := projectedControlByType(t, before, "MultiPage")
				if multiPage.SelectedIndex == nil || *multiPage.SelectedIndex != -1 {
					t.Fatalf("empty MultiPage selectedIndex = %#v, want -1", multiPage.SelectedIndex)
				}
				if got := projectedChildrenOfType(before, multiPage.ID, "Page"); got != 0 {
					t.Fatalf("empty MultiPage Page count = %d, want 0", got)
				}
			}

			original, err := oforms.SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			desired := templateInput(t, base, before)
			inputBefore := snapshotCopy(t, desired)
			compiled, err := CompileTemplate(base, desired, 932)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(inputBefore, desired) {
				t.Fatal("CompileTemplate mutated its input")
			}
			unchangedBase, err := oforms.SerializeForm(base, 932)
			if err != nil || !reflect.DeepEqual(original, unchangedBase) {
				t.Fatal("CompileTemplate mutated the source Form")
			}
			compiledBytes, err := oforms.SerializeForm(compiled, 932)
			if err != nil {
				t.Fatal(err)
			}
			assertDesignerBytes(t, original, compiledBytes)

			projected, err := projection.Project(compiled)
			if err != nil {
				t.Fatal(err)
			}
			if fixtureName == "baseline" {
				assertAuthoredPageMetadata(t, projected)
				assertAuthoredTabMetadata(t, projected)
			} else {
				multiPage := projectedControlByType(t, projected, "MultiPage")
				if multiPage.SelectedIndex == nil || *multiPage.SelectedIndex != -1 || projectedChildrenOfType(projected, multiPage.ID, "Page") != 0 {
					t.Fatalf("empty MultiPage projection = %#v", multiPage)
				}
			}
		})
	}
}

func TestCompileNewGeneratesMultiPageAndTabStripMetadata(t *testing.T) {
	input := newSpec()
	input.Controls = []spec.FormSpecControl{
		{ID: "pages", Name: "Pages", Type: "MultiPage", SelectedIndex: new(1), ZIndex: new(0)},
		{ID: "page-alpha", ParentID: "pages", Name: "PageAlpha", Type: "Page", Caption: new("Alpha"), ControlTipText: new("page-tip"), Tag: new("page-tag"), Accelerator: new("A"), Enabled: new(true), Visible: new(true), ZIndex: new(0)},
		{ID: "page-beta", ParentID: "pages", Name: "PageBeta", Type: "Page", Caption: new("Beta"), ControlTipText: new("beta-tip"), Tag: new("beta-tag"), Accelerator: new("B"), Enabled: new(true), Visible: new(true), ZIndex: new(1)},
		{ID: "tabs", Name: "Navigation", Type: "TabStrip", SelectedIndex: new(0), Tabs: []spec.FormSpecTab{{Name: "TabAlpha", Caption: new("Alpha"), ControlTipText: new("tab-tip"), Tag: new("tab-tag"), Accelerator: new("T"), Enabled: new(true), Visible: new(true)}}, ZIndex: new(1)},
		{ID: "empty-pages", Name: "EmptyPages", Type: "MultiPage", SelectedIndex: new(-1), ZIndex: new(2)},
		{ID: "empty-tabs", Name: "EmptyTabs", Type: "TabStrip", SelectedIndex: new(-1), Tabs: []spec.FormSpecTab{}, ZIndex: new(3)},
	}

	generated, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	projected, err := projection.Project(generated)
	if err != nil {
		t.Fatal(err)
	}
	pages := projectedControlByName(t, projected, "Pages")
	if pages.SelectedIndex == nil || *pages.SelectedIndex != 1 || projectedChildrenOfType(projected, pages.ID, "Page") != 2 {
		t.Fatalf("generated MultiPage = %#v", pages)
	}
	for _, want := range []struct {
		name, caption, tip, tag, accelerator string
	}{
		{"PageAlpha", "Alpha", "page-tip", "page-tag", "A"},
		{"PageBeta", "Beta", "beta-tip", "beta-tag", "B"},
	} {
		page := projectedControlByName(t, projected, want.name)
		if page.Type != "Page" || page.ParentID != pages.ID || page.Caption == nil || *page.Caption != want.caption || page.ControlTipText == nil || *page.ControlTipText != want.tip || page.Tag == nil || *page.Tag != want.tag || page.Accelerator == nil || *page.Accelerator != want.accelerator {
			t.Fatalf("generated Page metadata = %#v, want %+v", page, want)
		}
	}
	tabs := projectedControlByName(t, projected, "Navigation")
	if tabs.Type != "TabStrip" || tabs.SelectedIndex == nil || *tabs.SelectedIndex != 0 || len(tabs.Tabs) != 1 {
		t.Fatalf("generated TabStrip = %#v", tabs)
	}
	tab := tabs.Tabs[0]
	if tab.Name != "TabAlpha" || tab.Caption == nil || *tab.Caption != "Alpha" || tab.ControlTipText == nil || *tab.ControlTipText != "tab-tip" || tab.Tag == nil || *tab.Tag != "tab-tag" || tab.Accelerator == nil || *tab.Accelerator != "T" || tab.Enabled == nil || !*tab.Enabled || tab.Visible == nil || !*tab.Visible {
		t.Fatalf("generated tab metadata = %#v", tab)
	}
	emptyPages := projectedControlByName(t, projected, "EmptyPages")
	if emptyPages.SelectedIndex == nil || *emptyPages.SelectedIndex != -1 || projectedChildrenOfType(projected, emptyPages.ID, "Page") != 0 {
		t.Fatalf("generated empty MultiPage = %#v", emptyPages)
	}
	emptyTabs := projectedControlByName(t, projected, "EmptyTabs")
	if emptyTabs.SelectedIndex == nil || *emptyTabs.SelectedIndex != -1 || emptyTabs.Tabs == nil || len(emptyTabs.Tabs) != 0 {
		t.Fatalf("generated empty TabStrip = %#v", emptyTabs)
	}
}

func TestCompileTemplateAndEditsUpdatePageTopLevelMetadata(t *testing.T) {
	for _, mode := range []string{"template", "edits"} {
		t.Run(mode, func(t *testing.T) {
			base, before := multipageExcelAuthoredFixture(t, "baseline")
			desired := snapshotCopy(t, before)
			page := projectedControlByName(t, desired, "PageAlpha")
			page.Caption = new("Edited alpha")
			page.ControlTipText = new("edited page tip")
			page.Tag = new("edited page tag")
			page.Accelerator = new("E")
			inputBefore := snapshotCopy(t, desired)

			var result *oforms.Form
			var err error
			if mode == "template" {
				result, err = CompileTemplate(base, desired, 932)
			} else {
				result, err = CompileEdits(base, before, desired, 932)
			}
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(inputBefore, desired) {
				t.Fatalf("Compile%s mutated its after input", mode)
			}
			projected, err := projection.Project(result)
			if err != nil {
				t.Fatal(err)
			}
			updated := projectedControlByName(t, projected, "PageAlpha")
			if updated.Caption == nil || *updated.Caption != "Edited alpha" || updated.ControlTipText == nil || *updated.ControlTipText != "edited page tip" || updated.Tag == nil || *updated.Tag != "edited page tag" || updated.Accelerator == nil || *updated.Accelerator != "E" {
				t.Fatalf("edited Page top-level metadata = %#v", updated)
			}
		})
	}
}

func multipageExcelAuthoredFixture(t *testing.T, name string) (*oforms.Form, spec.FormSpec) {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("testdata", "multipage-excel-authored", name+".bin"))
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	names := oforms.DiscoverForms(container)
	if len(names) != 1 {
		t.Fatalf("Excel-authored fixture forms = %v, want one form", names)
	}
	form, err := oforms.ReadForm(container, names[0], 932)
	if err != nil {
		t.Fatal(err)
	}
	snapshot, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	return form, snapshot
}

func assertAuthoredPageMetadata(t *testing.T, snapshot spec.FormSpec) {
	t.Helper()
	page := projectedControlByName(t, snapshot, "PageAlpha")
	if page.Type != "Page" || page.Caption == nil || *page.Caption != "Alpha" || page.ControlTipText == nil || *page.ControlTipText != "alpha-tip" || page.Tag == nil || *page.Tag != "alpha-tag" || page.Accelerator == nil || *page.Accelerator != "A" || page.Enabled == nil || !*page.Enabled || page.Visible == nil || !*page.Visible {
		t.Fatalf("Excel-authored Page snapshot metadata = %#v", page)
	}
}

func assertAuthoredTabMetadata(t *testing.T, snapshot spec.FormSpec) {
	t.Helper()
	var owner *spec.FormSpecControl
	for i := range snapshot.Controls {
		control := &snapshot.Controls[i]
		if !strings.EqualFold(control.Type, "TabStrip") {
			continue
		}
		for _, tab := range control.Tabs {
			if tab.Name == "TabAlpha" {
				owner = control
				if tab.Caption == nil || *tab.Caption != "Alpha" || tab.ControlTipText == nil || *tab.ControlTipText != "tab-tip" || tab.Tag == nil || *tab.Tag != "tab-tag" || tab.Accelerator == nil || *tab.Accelerator != "T" || tab.Enabled == nil || !*tab.Enabled || tab.Visible == nil || !*tab.Visible {
					t.Fatalf("Excel-authored TabStrip tab metadata = %#v", tab)
				}
				break
			}
		}
	}
	if owner == nil {
		t.Fatal("Excel-authored snapshot has no TabStrip with TabAlpha")
	}
}

func projectedControlByType(t *testing.T, snapshot spec.FormSpec, typeName string) *spec.FormSpecControl {
	t.Helper()
	for i := range snapshot.Controls {
		if strings.EqualFold(snapshot.Controls[i].Type, typeName) {
			return &snapshot.Controls[i]
		}
	}
	t.Fatalf("projected snapshot has no %s control", typeName)
	return nil
}

func projectedControlByName(t *testing.T, snapshot spec.FormSpec, name string) *spec.FormSpecControl {
	t.Helper()
	for i := range snapshot.Controls {
		if snapshot.Controls[i].Name == name {
			return &snapshot.Controls[i]
		}
	}
	t.Fatalf("projected snapshot has no control named %q", name)
	return nil
}

func projectedChildrenOfType(snapshot spec.FormSpec, parentID, typeName string) int {
	count := 0
	for _, control := range snapshot.Controls {
		if control.ParentID == parentID && strings.EqualFold(control.Type, typeName) {
			count++
		}
	}
	return count
}

func assertRelocatedDesignerSubtreeBytes(t *testing.T, before, after *oforms.SerializedForm, oldPath, newPath string) {
	t.Helper()
	oldPrefix, newPrefix := oldPath+"/", newPath+"/"
	oldCount, newCount := 0, 0
	for path, raw := range before.Streams {
		if !strings.HasPrefix(path, oldPrefix) {
			continue
		}
		oldCount++
		newPath := newPrefix + strings.TrimPrefix(path, oldPrefix)
		if actual, ok := after.Streams[newPath]; !ok || !bytes.Equal(raw, actual) {
			t.Fatalf("relocated MultiPage subtree stream %q was not byte-preserved at %q", path, newPath)
		}
	}
	for path := range after.Streams {
		if strings.HasPrefix(path, newPrefix) {
			newCount++
		}
	}
	if oldCount == 0 || oldCount != newCount {
		t.Fatalf("relocated MultiPage subtree stream counts = %d -> %d", oldCount, newCount)
	}
}
