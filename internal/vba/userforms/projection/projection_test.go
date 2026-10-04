package projection

import (
	"encoding/json"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestProjectSimpleFixture(t *testing.T) {
	form := readFixture(t, "p4_form.bin")
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if got.SchemaVersion != 1 || got.Kind != "xlflow.userform" || got.Basis != "designer" {
		t.Fatalf("document identity = %#v", got)
	}
	if got.CoordinateSystem != "parent-relative" || got.Form.Name != "UserForm1" {
		t.Fatalf("form identity = %#v", got.Form)
	}
	if got.Form.Width != nil || got.Form.Height != nil {
		t.Fatalf("legacy form dimensions should remain absent, got %#v/%#v", got.Form.Width, got.Form.Height)
	}
	if got.Form.Observed == nil || got.Form.Observed.Width != nil || got.Form.Observed.Height != nil || got.Form.Observed.InsideWidth != nil || got.Form.Observed.InsideHeight != nil || got.Form.Observed.ClientWidth == nil || got.Form.Observed.ClientHeight == nil || *got.Form.Observed.ClientWidth <= 0 || *got.Form.Observed.ClientHeight <= 0 {
		t.Fatalf("observed form dimensions = %#v", got.Form.Observed)
	}
	if got.Form.Build == nil || got.Form.Build.Width != nil || got.Form.Build.Height != nil || got.Form.Build.ClientWidth == nil || got.Form.Build.ClientHeight == nil || *got.Form.Build.ClientWidth <= 0 || *got.Form.Build.ClientHeight <= 0 {
		t.Fatalf("build form dimensions = %#v", got.Form.Build)
	}
	if len(got.Controls) == 0 {
		t.Fatal("projected form has no controls")
	}
	button := got.Controls[0]
	if button.ID != fmt.Sprintf("control_%03d", 1) || button.Name != "CommandButton1" || button.Type != "CommandButton" || button.ProgID != "Forms.CommandButton.1" {
		t.Fatalf("first control = %#v", button)
	}
	if button.Caption == nil || *button.Caption != "CommandButton1" {
		t.Fatalf("button caption = %#v", button.Caption)
	}
	if button.Width == nil || math.Abs(*button.Width-126) > 0.001 {
		t.Fatalf("button width = %#v, want 126 points", button.Width)
	}
	wantHeight := 2328 * 72.0 / 2540.0
	if button.Height == nil || math.Abs(*button.Height-wantHeight) > 0.001 {
		t.Fatalf("button height = %#v, want %v points", button.Height, wantHeight)
	}
	if button.Visible == nil || !*button.Visible || button.Enabled == nil || !*button.Enabled {
		t.Fatalf("button visible/enabled = %#v/%#v", button.Visible, button.Enabled)
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("projected fixture did not validate: %#v", issues)
	}
}

func TestProjectControlCaptionOnlyWhenControlContractAllowsIt(t *testing.T) {
	tests := []struct {
		name        string
		controlType string
		caption     string
		wantCaption bool
	}{
		{name: "label", controlType: "Label", caption: "Title", wantCaption: true},
		{name: "page", controlType: "Page", caption: "Details", wantCaption: true},
		{name: "multipage-empty-caption", controlType: "MultiPage", caption: "", wantCaption: false},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			projected := forms.FormSpecControl{
				ID:   "controlNode",
				Type: test.controlType,
				Name: "Control1",
			}
			projectControl(&oforms.Control{
				Record: &oforms.Record{
					Strings: map[string]oforms.StoredString{
						"Caption": {Text: test.caption},
					},
				},
			}, &projected)
			if (projected.Caption != nil) != test.wantCaption {
				t.Fatalf("Caption = %#v, want present=%t", projected.Caption, test.wantCaption)
			}
			controls := []forms.FormSpecControl{projected}
			if test.controlType == "Page" {
				projected.ID = "pageNode"
				projected.ParentID = "multiNode"
				controls = []forms.FormSpecControl{
					{ID: "multiNode", Type: "MultiPage", Name: "PagesMain"},
					projected,
				}
			}
			form := forms.FormSpec{
				SchemaVersion: 1,
				Kind:          "xlflow.userform",
				Basis:         "designer",
				Form:          forms.FormSpecForm{Name: "ProjectionForm"},
				Controls:      controls,
			}
			if issues := forms.ValidateFormSpecStrict(form); hasErrors(issues) {
				t.Fatalf("projected control violates its contract: %#v", issues)
			}
		})
	}
}

func TestProjectControlCaptionReportingMatchesControlContract(t *testing.T) {
	tests := []struct {
		name            string
		controlType     string
		caption         string
		wantCaption     bool
		wantUnsupported bool
	}{
		{name: "label", controlType: "Label", caption: "Title", wantCaption: true},
		{name: "label-empty", controlType: "Label", wantCaption: true},
		{name: "frame", controlType: "Frame", caption: "Group", wantCaption: true},
		{name: "textbox", controlType: "TextBox", caption: "Internal label", wantUnsupported: true},
		{name: "textbox-empty", controlType: "TextBox"},
		{name: "multipage", controlType: "MultiPage", caption: "Internal pages", wantUnsupported: true},
		{name: "multipage-empty", controlType: "MultiPage"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			record := emptyRecord(test.controlType)
			record.Strings["Caption"] = oforms.StoredString{Text: test.caption}
			form := &oforms.Form{
				Name: "CaptionForm",
				Levels: []*oforms.Level{{
					Record: emptyRecord("Form"),
				}},
				Controls: []*oforms.Control{{
					Name: "ProjectedInput", Kind: "MSForms." + test.controlType, Record: record,
				}},
			}
			got, err := Project(form)
			if err != nil {
				t.Fatal(err)
			}
			if len(got.Controls) != 1 {
				t.Fatalf("controls = %d, want 1", len(got.Controls))
			}
			control := got.Controls[0]
			if (control.Caption != nil) != test.wantCaption {
				t.Fatalf("Caption = %#v, want present=%t", control.Caption, test.wantCaption)
			}
			if control.Caption != nil && *control.Caption != test.caption {
				t.Fatalf("Caption = %q, want %q", *control.Caption, test.caption)
			}
			var wantUnsupported []string
			if test.wantUnsupported {
				wantUnsupported = []string{"caption"}
			}
			if !slices.Equal(control.Unsupported, wantUnsupported) {
				t.Fatalf("unsupported = %q, want %q", control.Unsupported, wantUnsupported)
			}
			if test.wantUnsupported {
				if len(got.Warnings) != 1 || got.Warnings[0].Code != "unsupported_properties" ||
					got.Warnings[0].Control != control.Name || !strings.Contains(got.Warnings[0].Message, "caption") {
					t.Fatalf("warnings = %#v, want control caption warning", got.Warnings)
				}
			} else if len(got.Warnings) != 0 {
				t.Fatalf("warnings = %#v, want none", got.Warnings)
			}
			if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
				t.Fatalf("projected control violates its contract: %#v", issues)
			}
			if record.Strings["Caption"].Text != test.caption {
				t.Fatal("projection changed the persisted caption")
			}
		})
	}
}

func TestProjectNestedFixturePreservesParentsAndSiblingOrder(t *testing.T) {
	form := readFixture(t, "p6_nested_form.bin")
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	byID := make(map[string]forms.FormSpecControl, len(got.Controls))
	children := 0
	for _, control := range got.Controls {
		byID[control.ID] = control
		if control.ParentID != "" {
			children++
			if _, ok := byID[control.ParentID]; !ok {
				t.Fatalf("control %q appeared before parent %q", control.ID, control.ParentID)
			}
		}
	}
	if children == 0 {
		t.Fatal("nested fixture projected no parent relationships")
	}
	containerTypes := []string{"Frame", "MultiPage", "Page"}
	for _, control := range got.Controls {
		if control.Type == "TabStrip" && strings.HasPrefix(control.Name, "<unnamed_") {
			t.Fatalf("projected unnamed internal TabStrip: %#v", control)
		}
		if control.Enabled == nil {
			t.Fatalf("control %q omitted the persisted default enabled state", control.Name)
		}
		if control.ParentID == "" {
			continue
		}
		parent := byID[control.ParentID]
		if !slices.Contains(containerTypes, parent.Type) && !forms.IsUnsupportedControlPlaceholder(parent) {
			t.Fatalf("control %q has unexpected parent type %q", control.Name, parent.Type)
		}
	}
	for _, source := range form.Controls {
		if source == nil || source.MultiPage == nil {
			continue
		}
		projected, ok := controlByName(got.Controls, source.Name)
		if !ok {
			t.Fatalf("projected MultiPage %q is missing", source.Name)
		}
		if projected.Tabs != nil || projected.Observed != nil && projected.Observed.Tabs != nil {
			t.Fatalf("MultiPage %q exposed standalone tabs: %#v", source.Name, projected)
		}
		var gotPages []string
		for _, control := range got.Controls {
			if control.ParentID == projected.ID && control.Type == "Page" {
				gotPages = append(gotPages, control.Name)
			}
		}
		wantPages := make([]string, 0, len(source.MultiPage.Pages))
		for _, page := range source.MultiPage.Pages {
			wantPages = append(wantPages, page.Name)
		}
		if !slices.Equal(gotPages, wantPages) {
			t.Fatalf("MultiPage %q page order = %q, want x order %q", source.Name, gotPages, wantPages)
		}
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("projected nested fixture did not validate: %#v", issues)
	}
	path := filepath.Join(t.TempDir(), "NestedForm.json")
	if err := forms.WriteSnapshot(forms.SnapshotOutput{Path: path, DisplayPath: path, Format: "json"}, got); err != nil {
		t.Fatalf("WriteSnapshot: %v", err)
	}
	if _, err := forms.LoadFormSpec(forms.SpecInput{Path: path, DisplayPath: path, Format: "json"}); err != nil {
		t.Fatalf("LoadFormSpec: %v", err)
	}
}

func TestProjectOpaqueControlUsesSnapshotOnlyPlaceholder(t *testing.T) {
	form := &oforms.Form{
		Name: "OpaqueForm",
		Levels: []*oforms.Level{{
			Record: &oforms.Record{
				Type: "Form", Values: map[string]int64{}, Strings: map[string]oforms.StoredString{},
				Sizes: map[string]oforms.Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{},
			},
		}},
		Controls: []*oforms.Control{{
			Name: "VendorControl1", Kind: "ActiveX.Control", OpaqueRaw: []byte{1, 2, 3},
			Site: &oforms.Site{
				Values: map[string]int64{"BitFlags": 0x33}, Strings: map[string]oforms.StoredString{},
				Position: &oforms.Position{Left: 2540, Top: -1270},
			},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	control := got.Controls[0]
	if !forms.IsUnsupportedControlPlaceholder(control) {
		t.Fatalf("control = %#v, want snapshot placeholder", control)
	}
	if control.Left == nil || *control.Left != 72 || control.Top == nil || *control.Top != -36 {
		t.Fatalf("placeholder geometry = left %#v top %#v", control.Left, control.Top)
	}
	if !slices.Equal(control.Unsupported, []string{"controlType", "opaqueControlData"}) {
		t.Fatalf("unsupported = %q", control.Unsupported)
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("snapshot placeholder should validate with warnings: %#v", issues)
	}
	input := forms.SpecInput{DisplayPath: "OpaqueForm.yaml", Format: "yaml"}
	if err := forms.ValidateFormSpecForAuthoring(input, got); err == nil {
		t.Fatal("snapshot placeholder was accepted for authoring")
	}
}

func TestProjectDoesNotExposeUnsupportedCommandButtonValue(t *testing.T) {
	form := &oforms.Form{
		Name:   "ButtonForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "CommandButton1", Kind: "MSForms.CommandButton",
			Record: &oforms.Record{
				Type: "CommandButton", Values: map[string]int64{},
				Strings: map[string]oforms.StoredString{"Value": {Text: "False"}},
				Sizes:   map[string]oforms.Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{},
			},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	control := got.Controls[0]
	if control.Value != nil {
		t.Fatalf("CommandButton value = %#v, want omitted", control.Value)
	}
	if !slices.Contains(control.Unsupported, "value") {
		t.Fatalf("unsupported = %q, want value", control.Unsupported)
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("projected CommandButton did not validate: %#v", issues)
	}
}

func TestProjectNewControlValuesAndDefaultFlags(t *testing.T) {
	spin := emptyRecord("SpinButton")
	spin.Values["Position"] = 7
	spin.Values["VariousPropertyBits"] = 0x1b
	scroll := emptyRecord("ScrollBar")
	scroll.Values["Position"] = 42
	scroll.Values["VariousPropertyBits"] = 0x1b
	toggle := emptyRecord("ToggleButton")
	toggle.Strings["Value"] = oforms.StoredString{Text: "1"}
	toggle.Values["VariousPropertyBits"] = 0x2c80081b
	image := emptyRecord("Image")
	image.Values["VariousPropertyBits"] = 0x1b
	form := &oforms.Form{
		Name:   "NewControlsForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{
			{Name: "ToggleButton1", Kind: "MSForms.ToggleButton", Record: toggle},
			{Name: "SpinButton1", Kind: "MSForms.SpinButton", Record: spin},
			{Name: "ScrollBar1", Kind: "MSForms.ScrollBar", Record: scroll},
			{Name: "Image1", Kind: "MSForms.Image", Record: image},
		},
	}

	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	byName := make(map[string]forms.FormSpecControl, len(got.Controls))
	for _, control := range got.Controls {
		byName[control.Name] = control
	}
	if value := byName["ToggleButton1"].Value; value != "True" {
		t.Fatalf("ToggleButton value = %#v, want True", value)
	}
	for _, test := range []struct {
		name string
		want int
	}{
		{name: "SpinButton1", want: 7},
		{name: "ScrollBar1", want: 42},
	} {
		control := byName[test.name]
		if control.Value != test.want {
			t.Fatalf("%s value = %#v, want %d", test.name, control.Value, test.want)
		}
		if control.Enabled == nil || !*control.Enabled {
			t.Fatalf("%s enabled = %#v, want true", test.name, control.Enabled)
		}
		if slices.Contains(control.Unsupported, "position") || slices.Contains(control.Unsupported, "variousPropertyBits") {
			t.Fatalf("%s unsupported = %q, want Position and default bits projected", test.name, control.Unsupported)
		}
	}
	if imageControl := byName["Image1"]; imageControl.Enabled == nil || !*imageControl.Enabled || slices.Contains(imageControl.Unsupported, "variousPropertyBits") {
		t.Fatalf("Image projection = %#v, want enabled with default bits", imageControl)
	}
}

func TestProjectTabStripSnapshotCanBeWrittenAndLoaded(t *testing.T) {
	record := emptyRecord("TabStrip")
	record.Major = 2
	if err := oforms.SetTabStrip(record, &oforms.TabStrip{
		SelectedIndex: 1,
		Tabs: []oforms.Tab{
			{Name: "TabAlpha", Caption: "Alpha", ControlTipText: "alpha-tip", Tag: "alpha-tag", Accelerator: "A", Enabled: true, Visible: true},
			{Name: "TabBeta", Caption: "Beta", Enabled: false, Visible: false},
		},
	}); err != nil {
		t.Fatalf("SetTabStrip: %v", err)
	}
	form := &oforms.Form{
		Name:   "TabForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TabStrip1", Kind: "MSForms.TabStrip",
			Record: record,
		}},
	}
	projected, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if projected.Controls[0].SelectedIndex == nil || *projected.Controls[0].SelectedIndex != 1 {
		t.Fatalf("TabStrip selectedIndex = %#v, want 1", projected.Controls[0].SelectedIndex)
	}
	if len(projected.Controls[0].Tabs) != 2 || projected.Controls[0].Tabs[0].Name != "TabAlpha" || projected.Controls[0].Observed == nil || len(projected.Controls[0].Observed.Tabs) != 2 {
		t.Fatalf("TabStrip tabs = %#v, observed=%#v", projected.Controls[0].Tabs, projected.Controls[0].Observed)
	}
	for _, property := range []string{"selectedIndex", "items", "tipStrings", "tabNames", "tags", "accelerators"} {
		if slices.Contains(projected.Controls[0].Unsupported, property) {
			t.Fatalf("TabStrip unsupported = %q, includes projected property %q", projected.Controls[0].Unsupported, property)
		}
	}

	path := filepath.Join(t.TempDir(), "TabForm.json")
	if err := forms.WriteSnapshot(forms.SnapshotOutput{Path: path, DisplayPath: path, Format: "json"}, projected); err != nil {
		t.Fatalf("WriteSnapshot: %v", err)
	}
	loaded, err := forms.LoadFormSpec(forms.SpecInput{Path: path, DisplayPath: path, Format: "json"})
	if err != nil {
		t.Fatalf("LoadFormSpec: %v", err)
	}
	if loaded.Controls[0].SelectedIndex == nil || *loaded.Controls[0].SelectedIndex != 1 || len(loaded.Controls[0].Tabs) != 2 || loaded.Controls[0].Observed == nil || len(loaded.Controls[0].Observed.Tabs) != 2 {
		t.Fatalf("loaded TabStrip = %#v, want selected index and both tab slices", loaded.Controls[0])
	}
	if loaded.Controls[0].Tabs[0].ControlTipText == nil || *loaded.Controls[0].Tabs[0].ControlTipText != "alpha-tip" || loaded.Controls[0].Observed.Tabs[1].Visible == nil || *loaded.Controls[0].Observed.Tabs[1].Visible {
		t.Fatalf("loaded tab metadata = %#v, observed=%#v", loaded.Controls[0].Tabs, loaded.Controls[0].Observed.Tabs)
	}
}

func TestProjectEmptyTabStripNormalizesSelection(t *testing.T) {
	record := emptyRecord("TabStrip")
	record.Major = 2
	form := &oforms.Form{
		Name:   "EmptyTabForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TabStrip1", Kind: "MSForms.TabStrip", Record: record,
		}},
	}

	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	control := got.Controls[0]
	if control.SelectedIndex == nil || *control.SelectedIndex != -1 || control.Tabs == nil || len(control.Tabs) != 0 || control.Observed == nil || control.Observed.Tabs == nil || len(control.Observed.Tabs) != 0 {
		t.Fatalf("empty TabStrip projection = %#v, want empty tab slices and selectedIndex -1", control)
	}

	path := filepath.Join(t.TempDir(), "EmptyTabForm.json")
	if err := forms.WriteSnapshot(forms.SnapshotOutput{Path: path, DisplayPath: path, Format: "json"}, got); err != nil {
		t.Fatalf("WriteSnapshot: %v", err)
	}
	loaded, err := forms.LoadFormSpec(forms.SpecInput{Path: path, DisplayPath: path, Format: "json"})
	if err != nil {
		t.Fatalf("LoadFormSpec: %v", err)
	}
	if loaded.Controls[0].SelectedIndex == nil || *loaded.Controls[0].SelectedIndex != -1 || loaded.Controls[0].Tabs == nil || len(loaded.Controls[0].Tabs) != 0 {
		t.Fatalf("loaded empty TabStrip = %#v, want explicit empty tabs and selectedIndex -1", loaded.Controls[0])
	}
}

func TestProjectMultiPageAuthoredBaseline(t *testing.T) {
	for _, test := range []struct {
		name          string
		wantPageCount int
		wantSelected  int
	}{
		{name: "baseline.bin", wantPageCount: 2, wantSelected: 1},
		{name: "empty.bin", wantPageCount: 0, wantSelected: -1},
	} {
		t.Run(test.name, func(t *testing.T) {
			body, err := os.ReadFile(filepath.Join("..", "compiler", "testdata", "multipage-excel-authored", test.name))
			if err != nil {
				t.Fatal(err)
			}
			project, err := vbaproject.Read(body)
			if err != nil {
				t.Fatalf("read VBA project: %v", err)
			}
			if len(project.Forms) != 1 {
				t.Fatalf("UserForms = %d, want 1", len(project.Forms))
			}
			form := project.Forms[0]
			got, err := Project(form)
			if err != nil {
				t.Fatal(err)
			}
			var source *oforms.Control
			for _, control := range form.Controls {
				if control != nil && control.MultiPage != nil {
					source = control
					break
				}
			}
			if source == nil {
				t.Fatal("evidence form has no root MultiPage")
			}
			multiPage, ok := controlByName(got.Controls, source.Name)
			if !ok {
				t.Fatalf("projected MultiPage %q is missing", source.Name)
			}
			if multiPage.Tabs != nil || multiPage.Observed != nil && multiPage.Observed.Tabs != nil {
				t.Fatalf("MultiPage exposed TabStrip tabs: %#v", multiPage)
			}
			if multiPage.SelectedIndex == nil || *multiPage.SelectedIndex != test.wantSelected {
				t.Fatalf("MultiPage selectedIndex = %#v, want %d", multiPage.SelectedIndex, test.wantSelected)
			}
			var pages []forms.FormSpecControl
			for _, control := range got.Controls {
				if control.ParentID == multiPage.ID && control.Type == "Page" {
					pages = append(pages, control)
				}
			}
			if len(pages) != test.wantPageCount || len(source.MultiPage.Pages) != test.wantPageCount {
				t.Fatalf("projected/source pages = %d/%d, want %d", len(pages), len(source.MultiPage.Pages), test.wantPageCount)
			}
			if test.wantPageCount > 0 && (source.MultiPage.Hidden.TabStrip == nil || len(source.MultiPage.Hidden.TabStrip.Tabs) != test.wantPageCount) {
				t.Fatalf("hidden TabStrip = %#v, want one tab per page", source.MultiPage.Hidden.TabStrip)
			}
			if test.wantPageCount == 0 && (source.MultiPage.Hidden.TabStrip == nil || len(source.MultiPage.Hidden.TabStrip.Tabs) == 0) {
				t.Fatal("empty evidence should retain cached hidden tabs")
			}
			for index, page := range pages {
				want := source.MultiPage.Hidden.TabStrip.Tabs[index]
				if page.Name != source.MultiPage.Pages[index].Name || page.Caption == nil || *page.Caption != want.Caption {
					t.Errorf("page %d identity/caption = %#v, want %q / %q", index, page, source.MultiPage.Pages[index].Name, want.Caption)
				}
				if page.Enabled == nil || *page.Enabled != want.Enabled || page.Visible == nil || *page.Visible != want.Visible {
					t.Errorf("page %q enabled/visible = %#v/%#v, want %t/%t", page.Name, page.Enabled, page.Visible, want.Enabled, want.Visible)
				}
				if page.Left != nil || page.Top != nil || page.Width != nil || page.Height != nil {
					t.Errorf("page %q projected unavailable authoring geometry: %#v", page.Name, page)
				}
				if page.ControlTipText == nil || *page.ControlTipText != want.ControlTipText || page.Accelerator == nil || *page.Accelerator != want.Accelerator {
					t.Errorf("page %q top-level tab strings = tip %#v accelerator %#v, want %q / %q", page.Name, page.ControlTipText, page.Accelerator, want.ControlTipText, want.Accelerator)
				}
				if page.Properties != nil {
					for _, alias := range []string{"ControlTipText", "Accelerator"} {
						if _, exists := page.Properties[alias]; exists {
							t.Errorf("page %q duplicated %s in properties bag", page.Name, alias)
						}
					}
				}
				if expected, ok := source.MultiPage.Pages[index].Site.Strings["Tag"]; ok {
					if page.Tag == nil || *page.Tag != expected.Text {
						t.Errorf("page %q tag = %#v, want Site tag %q", page.Name, page.Tag, expected.Text)
					}
				}
			}
		})
	}
}

func TestProjectEmptyMultiPageNormalizesSelection(t *testing.T) {
	hiddenRecord := emptyRecord("TabStrip")
	hiddenRecord.Major = 2
	hiddenRecord.Mask = 1
	hiddenRecord.Values["ListIndex"] = 0
	multiPage := &oforms.Control{
		Name: "MultiPage1", Kind: "MSForms.MultiPage", Record: emptyRecord("MultiPage"),
		MultiPage: &oforms.MultiPage{
			Hidden:     &oforms.Control{Kind: "MSForms.TabStrip", Record: hiddenRecord},
			Properties: &oforms.Record{Type: "MultiPageProperties"},
		},
	}
	form := &oforms.Form{
		Name: "EmptyMultiPageForm", Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{multiPage},
	}

	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	control := got.Controls[0]
	if control.SelectedIndex == nil || *control.SelectedIndex != -1 || control.Tabs != nil || control.Observed != nil && control.Observed.Tabs != nil {
		t.Fatalf("empty MultiPage projection = %#v, want selectedIndex -1 and no tabs", control)
	}
}

func TestProjectExcelAuthoredDisabledPageUsesHiddenTabFlags(t *testing.T) {
	fixtureDir := filepath.Join("..", "compiler", "testdata", "multipage-excel-authored")
	body, err := os.ReadFile(filepath.Join(fixtureDir, "disabled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(body)
	if err != nil {
		t.Fatalf("read disabled fixture: %v", err)
	}
	if len(project.Forms) != 1 {
		t.Fatalf("UserForms = %d, want 1", len(project.Forms))
	}
	var source *oforms.Control
	for _, control := range project.Forms[0].Controls {
		if control != nil && control.MultiPage != nil {
			source = control
			break
		}
	}
	if source == nil || len(source.MultiPage.Pages) == 0 || source.MultiPage.Hidden.TabStrip == nil {
		t.Fatal("disabled fixture has no populated MultiPage")
	}
	pageRecordEnabled := source.MultiPage.Pages[0].Record.Values["BooleanProperties"]&4 != 0
	pageSiteVisible := source.MultiPage.Pages[0].Site.Values["BitFlags"]&(1<<1) != 0
	pageTab := source.MultiPage.Hidden.TabStrip.Tabs[0]
	if !pageRecordEnabled || pageSiteVisible || pageTab.Enabled || pageTab.Visible {
		t.Fatalf("fixture bits: Page BooleanProperties enabled=%t, Site visible=%t, hidden tab enabled/visible=%t/%t", pageRecordEnabled, pageSiteVisible, pageTab.Enabled, pageTab.Visible)
	}

	got, err := Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	page, ok := controlByName(got.Controls, source.MultiPage.Pages[0].Name)
	if !ok || page.Enabled == nil || *page.Enabled || page.Visible == nil || *page.Visible {
		t.Fatalf("projected disabled Page = %#v, want Enabled=false and Visible=false from hidden tab flags", page)
	}

	jsonBody, err := os.ReadFile(filepath.Join(fixtureDir, "disabled.json"))
	if err != nil {
		t.Fatal(err)
	}
	var snapshot struct {
		After struct {
			Pages []struct {
				Name    string `json:"name"`
				Enabled struct {
					Available bool `json:"available"`
					Value     bool `json:"value"`
				} `json:"Enabled"`
				Visible struct {
					Available bool `json:"available"`
					Value     bool `json:"value"`
				} `json:"Visible"`
			} `json:"pages"`
		} `json:"after"`
	}
	if err := json.Unmarshal(jsonBody, &snapshot); err != nil {
		t.Fatalf("decode Excel runtime snapshot: %v", err)
	}
	for _, captured := range snapshot.After.Pages {
		if captured.Name == source.MultiPage.Pages[0].Name && captured.Enabled.Available && captured.Visible.Available {
			if captured.Enabled.Value || captured.Visible.Value {
				t.Fatalf("Excel runtime reports PageAlpha enabled/visible=%t/%t, want false/false", captured.Enabled.Value, captured.Visible.Value)
			}
			return
		}
	}
	t.Fatal("Excel runtime snapshot does not contain available PageAlpha state")
}

func TestProjectReportsUnmodeledSiteStrings(t *testing.T) {
	form := &oforms.Form{
		Name:   "SiteForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TextBox1", Kind: "MSForms.TextBox",
			Site: &oforms.Site{Strings: map[string]oforms.StoredString{
				"Name":           {Text: "TextBox1"},
				"Tag":            {Text: "tag-value"},
				"ControlTipText": {Text: "tip"},
				"RuntimeLicKey":  {Text: "license"},
				"ControlSource":  {Text: "Sheet1!A1"},
				"RowSource":      {Text: "Sheet1!A1:A5"},
			}},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"controlSource", "controlTipText", "rowSource", "runtimeLicKey", "tag"}
	if !slices.Equal(got.Controls[0].Unsupported, want) {
		t.Fatalf("unsupported = %q, want %q", got.Controls[0].Unsupported, want)
	}
	warnings := 0
	for _, warning := range got.Warnings {
		if warning.Code == "unsupported_properties" && warning.Control == "TextBox1" {
			warnings++
			for _, name := range want {
				if !strings.Contains(warning.Message, name) {
					t.Errorf("warning %q omits %q", warning.Message, name)
				}
			}
		}
	}
	if warnings != 1 {
		t.Fatalf("unsupported_properties warnings for TextBox1 = %d, want 1", warnings)
	}
}

func TestProjectReportsUnmodeledControlFlags(t *testing.T) {
	form := &oforms.Form{
		Name:   "FlagsForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TextBox1", Kind: "MSForms.TextBox",
			Site:   &oforms.Site{Values: map[string]int64{"BitFlags": 0x35}},
			Record: &oforms.Record{Type: "TextBox", Values: map[string]int64{"VariousPropertyBits": 0x6}},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"siteFlags", "variousPropertyBits"}
	if !slices.Equal(got.Controls[0].Unsupported, want) {
		t.Fatalf("unsupported = %q, want %q", got.Controls[0].Unsupported, want)
	}
}

func TestProjectReportsNestedLevelAndFormBooleanProperties(t *testing.T) {
	nestedLevel := &oforms.Level{
		Record:       &oforms.Record{Type: "Form", Values: map[string]int64{"BooleanProperties": 2}},
		MouseIconRaw: []byte{1}, FontRaw: []byte{2}, PictureRaw: []byte{3},
		ExtraStreams: map[string][]byte{"custom": {4}},
	}
	form := &oforms.Form{
		Name: "NestedForm",
		Levels: []*oforms.Level{{Record: &oforms.Record{
			Type: "Form", Values: map[string]int64{"BooleanProperties": 1},
		}}},
		Controls: []*oforms.Control{{
			Name: "Frame1", Kind: "MSForms.Frame",
			Record: nestedLevel.Record, Level: nestedLevel,
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Warnings) != 2 {
		t.Fatalf("warnings = %#v, want form and Frame warnings", got.Warnings)
	}
	if !strings.Contains(got.Warnings[0].Message, "booleanProperties") {
		t.Fatalf("form warning = %#v, want booleanProperties", got.Warnings[0])
	}
	frame := got.Controls[0]
	want := []string{"booleanProperties", "extraStream:custom", "font", "mouseIcon", "picture"}
	if !slices.Equal(frame.Unsupported, want) {
		t.Fatalf("Frame unsupported = %q, want %q", frame.Unsupported, want)
	}
	if got.Warnings[1].Control != "Frame1" || !strings.Contains(got.Warnings[1].Message, "picture") {
		t.Fatalf("Frame warning = %#v", got.Warnings[1])
	}
}

func TestProjectDoesNotReportPersistedBitfieldDefaults(t *testing.T) {
	form := &oforms.Form{
		Name: "DefaultsForm",
		Levels: []*oforms.Level{{Record: &oforms.Record{
			Type: "Form", Values: map[string]int64{"BooleanProperties": 0x4004},
		}}},
		Controls: []*oforms.Control{
			{
				Name: "Frame1", Kind: "MSForms.Frame",
				Site:   &oforms.Site{Values: map[string]int64{"BitFlags": 0x40023}},
				Record: &oforms.Record{Type: "Frame", Values: map[string]int64{"BooleanProperties": 0x8004}},
			},
			{
				Name: "TextBox1", Kind: "MSForms.TextBox",
				Site:   &oforms.Site{Values: map[string]int64{"BitFlags": 0x33}},
				Record: &oforms.Record{Type: "TextBox", Values: map[string]int64{"VariousPropertyBits": 0x2c80481b}},
			},
		},
	}

	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Warnings) != 0 {
		t.Fatalf("warnings = %#v, want no warnings for persisted defaults", got.Warnings)
	}
	for _, control := range got.Controls {
		if len(control.Unsupported) != 0 {
			t.Errorf("%s unsupported = %q, want none for persisted defaults", control.Name, control.Unsupported)
		}
	}
}

func TestProjectImagePictureDisplayProperties(t *testing.T) {
	form := &oforms.Form{
		Name:   "ImageProperties",
		Levels: []*oforms.Level{{Record: &oforms.Record{Type: "Form", Values: map[string]int64{}}}},
		Controls: []*oforms.Control{{
			Name: "Image1", Kind: "MSForms.Image",
			Site: &oforms.Site{Values: map[string]int64{"BitFlags": 0x33}},
			Record: &oforms.Record{
				Type: "Image",
				Values: map[string]int64{
					"VariousPropertyBits": 0x1b,
					"PictureAlignment":    3,
					"PictureSizeMode":     3,
				},
				Pictures: map[string][]byte{"Picture": {0x01}},
			},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	image := got.Controls[0]
	if image.Properties["pictureAlignment"] != 3 || image.Properties["pictureSizeMode"] != 3 {
		t.Fatalf("Image properties = %#v, want alignment and size mode 3", image.Properties)
	}
	if !slices.Equal(image.Unsupported, []string{"picture"}) {
		t.Fatalf("Image unsupported = %q, want picture only", image.Unsupported)
	}
}

func TestProjectedControlValueMatchesExcelBooleanSpelling(t *testing.T) {
	for _, test := range []struct {
		persisted string
		want      string
	}{
		{persisted: "1", want: "True"},
		{persisted: "0", want: "False"},
		{persisted: "-1", want: "True"},
	} {
		if got := projectedControlValue("CheckBox", test.persisted); got != test.want {
			t.Fatalf("projectedControlValue(CheckBox, %q) = %#v, want %q", test.persisted, got, test.want)
		}
	}
	if got := projectedControlValue("TextBox", "1"); got != "1" {
		t.Fatalf("projectedControlValue(TextBox, 1) = %#v, want raw text", got)
	}
}

func TestProjectIsDeterministic(t *testing.T) {
	form := readFixture(t, "p6_nested_form.bin")
	first, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	second, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	firstJSON, err := json.Marshal(first)
	if err != nil {
		t.Fatal(err)
	}
	secondJSON, err := json.Marshal(second)
	if err != nil {
		t.Fatal(err)
	}
	if string(firstJSON) != string(secondJSON) {
		t.Fatalf("projection is unstable:\n%s\n%s", firstJSON, secondJSON)
	}
}

func readFixture(t *testing.T, name string) *oforms.Form {
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
	return form
}

func hasErrors(issues []forms.ValidationIssue) bool {
	for _, issue := range issues {
		if issue.Severity == forms.SeverityError {
			return true
		}
	}
	return false
}

func controlByName(controls []forms.FormSpecControl, name string) (forms.FormSpecControl, bool) {
	for _, control := range controls {
		if control.Name == name {
			return control, true
		}
	}
	return forms.FormSpecControl{}, false
}

func emptyRecord(recordType string) *oforms.Record {
	return &oforms.Record{
		Type: recordType, Values: map[string]int64{}, Strings: map[string]oforms.StoredString{},
		Sizes: map[string]oforms.Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{},
	}
}
