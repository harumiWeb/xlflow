package forms

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"gopkg.in/yaml.v3"
)

type SnapshotOutput struct {
	Path        string
	DisplayPath string
	Format      string
}

type SpecInput struct {
	Path        string
	DisplayPath string
	Format      string
}

type FormSpec struct {
	SchemaVersion    int               `json:"schemaVersion" yaml:"schemaVersion"`
	Kind             string            `json:"kind" yaml:"kind"`
	Basis            string            `json:"basis" yaml:"basis"`
	CoordinateSystem string            `json:"coordinateSystem,omitempty" yaml:"coordinateSystem,omitempty"`
	Form             FormSpecForm      `json:"form" yaml:"form"`
	Controls         []FormSpecControl `json:"controls" yaml:"controls"`
	Warnings         []FormSpecWarning `json:"warnings" yaml:"warnings"`
}

type FormSpecForm struct {
	Name     string                `json:"name" yaml:"name"`
	Caption  *string               `json:"caption,omitempty" yaml:"caption,omitempty"`
	Width    *float64              `json:"width,omitempty" yaml:"width,omitempty"`
	Height   *float64              `json:"height,omitempty" yaml:"height,omitempty"`
	Observed *FormSpecObservedForm `json:"observed,omitempty" yaml:"observed,omitempty"`
	Build    *FormSpecBuildForm    `json:"build,omitempty" yaml:"build,omitempty"`
}

type FormSpecObservedForm struct {
	Caption      *string  `json:"caption,omitempty" yaml:"caption,omitempty"`
	Width        *float64 `json:"width,omitempty" yaml:"width,omitempty"`
	Height       *float64 `json:"height,omitempty" yaml:"height,omitempty"`
	InsideWidth  *float64 `json:"insideWidth,omitempty" yaml:"insideWidth,omitempty"`
	InsideHeight *float64 `json:"insideHeight,omitempty" yaml:"insideHeight,omitempty"`
}

type FormSpecBuildForm struct {
	Caption *string  `json:"caption,omitempty" yaml:"caption,omitempty"`
	Width   *float64 `json:"width,omitempty" yaml:"width,omitempty"`
	Height  *float64 `json:"height,omitempty" yaml:"height,omitempty"`
}

type FormSpecControl struct {
	ID            string                   `json:"id,omitempty" yaml:"id,omitempty"`
	ParentID      string                   `json:"parentId,omitempty" yaml:"parentId,omitempty"`
	ZIndex        *int                     `json:"zIndex,omitempty" yaml:"zIndex,omitempty"`
	Type          string                   `json:"type" yaml:"type"`
	Name          string                   `json:"name" yaml:"name"`
	ProgID        string                   `json:"progId,omitempty" yaml:"progId,omitempty"`
	Caption       *string                  `json:"caption,omitempty" yaml:"caption,omitempty"`
	Text          *string                  `json:"text,omitempty" yaml:"text,omitempty"`
	Value         any                      `json:"value,omitempty" yaml:"value,omitempty"`
	Left          *float64                 `json:"left,omitempty" yaml:"left,omitempty"`
	Top           *float64                 `json:"top,omitempty" yaml:"top,omitempty"`
	Width         *float64                 `json:"width,omitempty" yaml:"width,omitempty"`
	Height        *float64                 `json:"height,omitempty" yaml:"height,omitempty"`
	TabIndex      *int                     `json:"tabIndex,omitempty" yaml:"tabIndex,omitempty"`
	SelectedIndex *int                     `json:"selectedIndex,omitempty" yaml:"selectedIndex,omitempty"`
	Enabled       *bool                    `json:"enabled,omitempty" yaml:"enabled,omitempty"`
	Visible       *bool                    `json:"visible,omitempty" yaml:"visible,omitempty"`
	List          []string                 `json:"list,omitempty" yaml:"list,omitempty"`
	Unsupported   []string                 `json:"unsupported,omitempty" yaml:"unsupported,omitempty"`
	Controls      []FormSpecControl        `json:"controls,omitempty" yaml:"controls,omitempty"`
	Properties    map[string]any           `json:"properties,omitempty" yaml:"properties,omitempty"`
	Observed      *FormSpecObservedControl `json:"observed,omitempty" yaml:"observed,omitempty"`
}

type FormSpecObservedControl struct {
	Caption       *string        `json:"caption,omitempty" yaml:"caption,omitempty"`
	Text          *string        `json:"text,omitempty" yaml:"text,omitempty"`
	Value         any            `json:"value,omitempty" yaml:"value,omitempty"`
	Left          *float64       `json:"left,omitempty" yaml:"left,omitempty"`
	Top           *float64       `json:"top,omitempty" yaml:"top,omitempty"`
	Width         *float64       `json:"width,omitempty" yaml:"width,omitempty"`
	Height        *float64       `json:"height,omitempty" yaml:"height,omitempty"`
	TabIndex      *int           `json:"tabIndex,omitempty" yaml:"tabIndex,omitempty"`
	SelectedIndex *int           `json:"selectedIndex,omitempty" yaml:"selectedIndex,omitempty"`
	Enabled       *bool          `json:"enabled,omitempty" yaml:"enabled,omitempty"`
	Visible       *bool          `json:"visible,omitempty" yaml:"visible,omitempty"`
	List          []string       `json:"list,omitempty" yaml:"list,omitempty"`
	Unsupported   []string       `json:"unsupported,omitempty" yaml:"unsupported,omitempty"`
	Properties    map[string]any `json:"properties,omitempty" yaml:"properties,omitempty"`
}

type FormSpecWarning struct {
	Code    string `json:"code,omitempty" yaml:"code,omitempty"`
	Message string `json:"message,omitempty" yaml:"message,omitempty"`
	Control string `json:"control,omitempty" yaml:"control,omitempty"`
}

type SpecError struct {
	Code       string
	Message    string
	Path       string
	Format     string
	Line       int
	Column     int
	Field      string
	Suggestion string
	Cause      error
}

func (e *SpecError) Error() string {
	if e == nil {
		return ""
	}
	return e.Message
}

func (e *SpecError) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Cause
}

var typeProgIDMap = map[string]string{
	"label":         "Forms.Label.1",
	"textbox":       "Forms.TextBox.1",
	"combobox":      "Forms.ComboBox.1",
	"listbox":       "Forms.ListBox.1",
	"commandbutton": "Forms.CommandButton.1",
	"checkbox":      "Forms.CheckBox.1",
	"optionbutton":  "Forms.OptionButton.1",
	"frame":         "Forms.Frame.1",
}

func ResolveSnapshotOutput(root, outPath string) (SnapshotOutput, error) {
	trimmed := strings.TrimSpace(outPath)
	if trimmed == "" {
		return SnapshotOutput{}, fmt.Errorf("--out is required")
	}
	resolved := trimmed
	if !filepath.IsAbs(resolved) {
		resolved = filepath.Join(root, resolved)
	}
	resolved = filepath.Clean(resolved)
	format, err := snapshotFormatFromPath(resolved)
	if err != nil {
		return SnapshotOutput{}, err
	}
	if info, statErr := os.Stat(resolved); statErr == nil && info.IsDir() {
		return SnapshotOutput{}, fmt.Errorf("output path %q is a directory", trimmed)
	} else if statErr != nil && !os.IsNotExist(statErr) {
		return SnapshotOutput{}, statErr
	}
	return SnapshotOutput{
		Path:        resolved,
		DisplayPath: filepath.ToSlash(relPath(root, resolved)),
		Format:      format,
	}, nil
}

func WriteSnapshot(output SnapshotOutput, spec FormSpec) error {
	if err := os.MkdirAll(filepath.Dir(output.Path), 0o755); err != nil {
		return err
	}
	var (
		body []byte
		err  error
	)
	switch output.Format {
	case "json":
		body, err = json.MarshalIndent(spec, "", "  ")
		if err == nil {
			body = append(body, '\n')
		}
	case "yaml":
		body, err = yaml.Marshal(spec)
	default:
		err = fmt.Errorf("unsupported snapshot format %q", output.Format)
	}
	if err != nil {
		return err
	}
	return os.WriteFile(output.Path, body, 0o644)
}

func ResolveSpecInput(root, specPath string) (SpecInput, error) {
	trimmed := strings.TrimSpace(specPath)
	if trimmed == "" {
		return SpecInput{}, fmt.Errorf("spec path is required")
	}
	resolved := trimmed
	if !filepath.IsAbs(resolved) {
		resolved = filepath.Join(root, resolved)
	}
	resolved = filepath.Clean(resolved)
	format, err := snapshotFormatFromPath(resolved)
	if err != nil {
		return SpecInput{}, fmt.Errorf("spec path must end with .json, .yaml, or .yml")
	}
	info, statErr := os.Stat(resolved)
	if statErr != nil {
		if os.IsNotExist(statErr) {
			return SpecInput{}, fmt.Errorf("spec file %q was not found", trimmed)
		}
		return SpecInput{}, statErr
	}
	if info.IsDir() {
		return SpecInput{}, fmt.Errorf("spec path %q is a directory", trimmed)
	}
	return SpecInput{
		Path:        resolved,
		DisplayPath: filepath.ToSlash(relPath(root, resolved)),
		Format:      format,
	}, nil
}

func LoadFormSpec(input SpecInput) (FormSpec, error) {
	body, err := os.ReadFile(input.Path)
	if err != nil {
		return FormSpec{}, err
	}
	var spec FormSpec
	switch input.Format {
	case "json":
		err = json.Unmarshal(body, &spec)
	case "yaml":
		err = yaml.Unmarshal(body, &spec)
	default:
		err = fmt.Errorf("unsupported form spec format %q", input.Format)
	}
	if err != nil {
		return FormSpec{}, newSpecParseError(input, body, err)
	}
	spec = NormalizeFormSpec(spec)
	if err := ValidateFormSpec(spec); err != nil {
		var specErr *SpecError
		if errors.As(err, &specErr) {
			if specErr.Path == "" {
				specErr.Path = input.DisplayPath
			}
			if specErr.Format == "" {
				specErr.Format = input.Format
			}
		}
		return FormSpec{}, err
	}
	return spec, nil
}

func NormalizeFormSpec(spec FormSpec) FormSpec {
	spec.Form = normalizeFormSpecForm(spec.Form)
	normalizedControls, _ := normalizeFormSpecControls(spec.Controls)
	spec.Controls = normalizedControls
	if spec.Warnings == nil {
		spec.Warnings = []FormSpecWarning{}
	}
	return spec
}

func ValidateFormSpec(spec FormSpec) error {
	if spec.SchemaVersion != 1 {
		return newSpecValidationError("spec_schema_invalid", "schemaVersion must be 1", "schemaVersion")
	}
	if spec.Kind != "xlflow.userform" {
		return newSpecValidationError("spec_schema_invalid", `kind must be "xlflow.userform"`, "kind")
	}
	if strings.TrimSpace(spec.Basis) != "designer" {
		return newSpecValidationError("spec_schema_invalid", `basis must be "designer"`, "basis")
	}
	if strings.TrimSpace(spec.Form.Name) == "" {
		return newSpecValidationError("spec_schema_invalid", "form.name is required", "form.name")
	}
	ids := make(map[string]struct{}, len(spec.Controls))
	for i, control := range spec.Controls {
		path := fmt.Sprintf("controls[%d]", i)
		if err := ValidateFormSpecControl(control, path); err != nil {
			return err
		}
		if _, exists := ids[control.ID]; exists {
			return newSpecValidationError("spec_validation_failed", fmt.Sprintf("%s.id %q is duplicated", path, control.ID), path+".id")
		}
		ids[control.ID] = struct{}{}
	}
	for i, control := range spec.Controls {
		if strings.TrimSpace(control.ParentID) == "" {
			continue
		}
		if _, ok := ids[control.ParentID]; !ok {
			field := fmt.Sprintf("controls[%d].parentId", i)
			return newSpecValidationError("spec_validation_failed", fmt.Sprintf("%s %q was not found", field, control.ParentID), field)
		}
	}
	return nil
}

func ValidateFormSpecControl(control FormSpecControl, path string) error {
	if strings.TrimSpace(control.ID) == "" {
		return newSpecValidationError("spec_validation_failed", path+".id is required", path+".id")
	}
	if strings.TrimSpace(control.Name) == "" {
		return newSpecValidationError("spec_validation_failed", path+".name is required", path+".name")
	}
	if strings.TrimSpace(control.Type) == "" {
		return newSpecValidationError("spec_validation_failed", path+".type is required", path+".type")
	}
	if _, err := ControlProgID(control); err != nil {
		return newSpecValidationError("spec_validation_failed", fmt.Sprintf("%s: %v", path, err), path+".type")
	}
	return nil
}

func ControlProgID(control FormSpecControl) (string, error) {
	if progID := strings.TrimSpace(control.ProgID); progID != "" {
		return progID, nil
	}
	progID, ok := typeProgIDMap[strings.ToLower(strings.TrimSpace(control.Type))]
	if !ok {
		return "", fmt.Errorf("unsupported control type %q", control.Type)
	}
	return progID, nil
}

func FormSpecFromInspectSnapshot(snapshot any) (FormSpec, error) {
	root, ok := asObjectMap(snapshot)
	if !ok {
		return FormSpec{}, fmt.Errorf("inspect designer snapshot payload is missing or invalid")
	}
	name, ok := stringField(root, "name")
	if !ok || strings.TrimSpace(name) == "" {
		return FormSpec{}, fmt.Errorf("inspect designer snapshot did not include a form name")
	}
	basis, _ := stringField(root, "basis")
	if basis == "" {
		basis = "designer"
	}
	coordinateSystem, _ := stringField(root, "coordinate_system")
	placeholderCounter := 0
	idCounter := 0
	controls, generatedWarnings, err := formSpecControls(root["controls"], "", &placeholderCounter, &idCounter)
	if err != nil {
		return FormSpec{}, err
	}
	warnings, err := formSpecWarnings(root["warnings"])
	if err != nil {
		return FormSpec{}, err
	}
	warnings = append(warnings, generatedWarnings...)
	form := FormSpecForm{Name: name}
	observed := &FormSpecObservedForm{}
	build := &FormSpecBuildForm{}
	if caption, ok := stringField(root, "caption"); ok {
		form.Caption = &caption
		observed.Caption = &caption
		build.Caption = &caption
	}
	if width, ok := optionalFloatField(root, "width"); ok {
		form.Width = &width
		observed.Width = &width
		build.Width = &width
	}
	if height, ok := optionalFloatField(root, "height"); ok {
		form.Height = &height
		observed.Height = &height
		build.Height = &height
	}
	if !hasObservedFormValues(observed) {
		observed = nil
	}
	if !hasBuildFormValues(build) {
		build = nil
	}
	form.Observed = observed
	form.Build = build
	if warnings == nil {
		warnings = []FormSpecWarning{}
	}
	return NormalizeFormSpec(FormSpec{
		SchemaVersion:    1,
		Kind:             "xlflow.userform",
		Basis:            basis,
		CoordinateSystem: coordinateSystem,
		Form:             form,
		Controls:         controls,
		Warnings:         warnings,
	}), nil
}

func normalizeFormSpecForm(form FormSpecForm) FormSpecForm {
	if form.Observed == nil {
		form.Observed = &FormSpecObservedForm{}
	}
	if form.Observed.Caption == nil && form.Caption != nil {
		form.Observed.Caption = form.Caption
	}
	if form.Observed.Width == nil && form.Width != nil {
		form.Observed.Width = form.Width
	}
	if form.Observed.Height == nil && form.Height != nil {
		form.Observed.Height = form.Height
	}
	if !hasObservedFormValues(form.Observed) {
		form.Observed = nil
	}
	if form.Build == nil {
		form.Build = &FormSpecBuildForm{}
	}
	if form.Build.Caption == nil {
		form.Build.Caption = firstStringPtr(form.Caption, observedFormCaption(form.Observed))
	}
	if form.Build.Width == nil {
		form.Build.Width = firstFloatPtr(form.Width, observedFormWidth(form.Observed))
	}
	if form.Build.Height == nil {
		form.Build.Height = firstFloatPtr(form.Height, observedFormHeight(form.Observed))
	}
	if !hasBuildFormValues(form.Build) {
		form.Build = nil
	}
	return form
}

func normalizeFormSpecControls(controls []FormSpecControl) ([]FormSpecControl, []FormSpecWarning) {
	state := &normalizeControlState{
		nextID:   1,
		usedIDs:  map[string]struct{}{},
		warnings: []FormSpecWarning{},
	}
	normalized := make([]FormSpecControl, 0)
	for index, control := range controls {
		normalized = append(normalized, state.normalizeControl(control, "", index)...)
	}
	return normalized, state.warnings
}

type normalizeControlState struct {
	nextID   int
	usedIDs  map[string]struct{}
	warnings []FormSpecWarning
}

func (s *normalizeControlState) normalizeControl(control FormSpecControl, inheritedParentID string, index int) []FormSpecControl {
	control = normalizeObservedControl(control)
	control.ID = s.normalizeControlID(control.ID, control.Name)
	if strings.TrimSpace(control.ParentID) == "" {
		control.ParentID = inheritedParentID
	}
	if control.ZIndex == nil {
		z := index
		control.ZIndex = &z
	}
	children := control.Controls
	control.Controls = nil
	items := []FormSpecControl{control}
	for childIndex, child := range children {
		items = append(items, s.normalizeControl(child, control.ID, childIndex)...)
	}
	return items
}

func (s *normalizeControlState) normalizeControlID(existingID, name string) string {
	id := strings.TrimSpace(existingID)
	if id == "" {
		return s.uniqueControlID(name)
	}
	if _, exists := s.usedIDs[id]; !exists {
		s.usedIDs[id] = struct{}{}
	}
	return id
}

func (s *normalizeControlState) uniqueControlID(name string) string {
	base := strings.TrimSpace(name)
	if base == "" {
		base = "control"
	}
	base = strings.ToLower(strings.ReplaceAll(base, " ", "_"))
	id := base
	if _, exists := s.usedIDs[id]; !exists {
		s.usedIDs[id] = struct{}{}
		return id
	}
	for {
		id = fmt.Sprintf("%s_%03d", base, s.nextID)
		s.nextID++
		if _, exists := s.usedIDs[id]; exists {
			continue
		}
		s.usedIDs[id] = struct{}{}
		return id
	}
}

func normalizeObservedControl(control FormSpecControl) FormSpecControl {
	if control.Observed == nil {
		control.Observed = &FormSpecObservedControl{}
	}
	if control.Observed.Caption == nil && control.Caption != nil {
		control.Observed.Caption = control.Caption
	}
	if control.Observed.Text == nil && control.Text != nil {
		control.Observed.Text = control.Text
	}
	if control.Observed.Value == nil && control.Value != nil {
		control.Observed.Value = control.Value
	}
	if control.Observed.Left == nil && control.Left != nil {
		control.Observed.Left = control.Left
	}
	if control.Observed.Top == nil && control.Top != nil {
		control.Observed.Top = control.Top
	}
	if control.Observed.Width == nil && control.Width != nil {
		control.Observed.Width = control.Width
	}
	if control.Observed.Height == nil && control.Height != nil {
		control.Observed.Height = control.Height
	}
	if control.Observed.TabIndex == nil && control.TabIndex != nil {
		control.Observed.TabIndex = control.TabIndex
	}
	if control.Observed.SelectedIndex == nil && control.SelectedIndex != nil {
		control.Observed.SelectedIndex = control.SelectedIndex
	}
	if control.Observed.Enabled == nil && control.Enabled != nil {
		control.Observed.Enabled = control.Enabled
	}
	if control.Observed.Visible == nil && control.Visible != nil {
		control.Observed.Visible = control.Visible
	}
	if len(control.Observed.List) == 0 && len(control.List) > 0 {
		control.Observed.List = append([]string(nil), control.List...)
	}
	if len(control.Observed.Unsupported) == 0 && len(control.Unsupported) > 0 {
		control.Observed.Unsupported = append([]string(nil), control.Unsupported...)
	}
	if len(control.Observed.Properties) == 0 && len(control.Properties) > 0 {
		control.Observed.Properties = cloneMap(control.Properties)
	}
	if !hasObservedControlValues(control.Observed) {
		control.Observed = nil
	}
	return control
}

func snapshotFormatFromPath(path string) (string, error) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".json":
		return "json", nil
	case ".yaml", ".yml":
		return "yaml", nil
	default:
		return "", fmt.Errorf("--out must end with .json, .yaml, or .yml")
	}
}

func formSpecControls(value any, parentID string, unnamedCounter *int, idCounter *int) ([]FormSpecControl, []FormSpecWarning, error) {
	items, ok := asSlice(value)
	if !ok || len(items) == 0 {
		return []FormSpecControl{}, []FormSpecWarning{}, nil
	}
	controls := make([]FormSpecControl, 0, len(items))
	warnings := make([]FormSpecWarning, 0)
	for index, item := range items {
		controlMap, ok := asObjectMap(item)
		if !ok {
			return nil, nil, fmt.Errorf("inspect designer snapshot control entry is invalid")
		}
		control, childControls, controlWarnings, err := formSpecControl(controlMap, parentID, index, unnamedCounter, idCounter)
		if err != nil {
			return nil, nil, err
		}
		controls = append(controls, control)
		controls = append(controls, childControls...)
		warnings = append(warnings, controlWarnings...)
	}
	return controls, warnings, nil
}

func formSpecControl(root map[string]any, parentID string, index int, unnamedCounter *int, idCounter *int) (FormSpecControl, []FormSpecControl, []FormSpecWarning, error) {
	name, ok := stringField(root, "name")
	warnings := make([]FormSpecWarning, 0)
	if !ok || strings.TrimSpace(name) == "" {
		if unnamedCounter == nil {
			return FormSpecControl{}, nil, nil, fmt.Errorf("inspect designer snapshot control is missing a name")
		}
		*unnamedCounter = *unnamedCounter + 1
		name = fmt.Sprintf("<unnamed_%d>", *unnamedCounter)
		warnings = append(warnings, FormSpecWarning{
			Code:    "unnamed_control_placeholder",
			Message: "A control without a stable name was persisted with a generated placeholder name.",
			Control: name,
		})
	}
	controlType, ok := stringField(root, "type")
	if !ok || strings.TrimSpace(controlType) == "" {
		return FormSpecControl{}, nil, nil, fmt.Errorf("inspect designer snapshot control %q is missing a type", name)
	}
	id, _ := stringField(root, "id")
	control := FormSpecControl{
		ID:       strings.TrimSpace(id),
		ParentID: strings.TrimSpace(parentID),
		Name:     name,
		Type:     controlType,
	}
	z := index
	control.ZIndex = &z
	if control.ID == "" {
		*idCounter = *idCounter + 1
		control.ID = fmt.Sprintf("control_%03d", *idCounter)
	}
	if progID, ok := stringField(root, "prog_id"); ok {
		control.ProgID = progID
	}
	if progID, ok := stringField(root, "progId"); ok && control.ProgID == "" {
		control.ProgID = progID
	}
	if caption, ok := stringField(root, "caption"); ok {
		control.Caption = &caption
	}
	if text, ok := stringField(root, "text"); ok {
		control.Text = &text
	}
	if value, ok := root["value"]; ok {
		control.Value = value
	}
	if left, ok := optionalFloatField(root, "left"); ok {
		control.Left = &left
	}
	if top, ok := optionalFloatField(root, "top"); ok {
		control.Top = &top
	}
	if width, ok := optionalFloatField(root, "width"); ok {
		control.Width = &width
	}
	if height, ok := optionalFloatField(root, "height"); ok {
		control.Height = &height
	}
	if tabIndex, ok := optionalIntField(root, "tab_index"); ok {
		control.TabIndex = &tabIndex
	}
	if selectedIndex, ok := optionalIntField(root, "selected_index"); ok {
		control.SelectedIndex = &selectedIndex
	}
	if enabled, ok := optionalBoolField(root, "enabled"); ok {
		control.Enabled = &enabled
	}
	if visible, ok := optionalBoolField(root, "visible"); ok {
		control.Visible = &visible
	}
	if list, ok := stringSliceField(root, "list"); ok {
		control.List = list
	}
	if unsupported, ok := stringSliceField(root, "unsupported"); ok {
		control.Unsupported = unsupported
	}
	if properties, ok := asObjectMap(root["properties"]); ok && len(properties) > 0 {
		control.Properties = properties
	}
	control = normalizeObservedControl(control)
	children, childWarnings, err := formSpecControls(root["controls"], control.ID, unnamedCounter, idCounter)
	if err != nil {
		return FormSpecControl{}, nil, nil, err
	}
	warnings = append(warnings, childWarnings...)
	return control, children, warnings, nil
}

func formSpecWarnings(value any) ([]FormSpecWarning, error) {
	items, ok := asSlice(value)
	if !ok || len(items) == 0 {
		return []FormSpecWarning{}, nil
	}
	warnings := make([]FormSpecWarning, 0, len(items))
	for _, item := range items {
		warningMap, ok := asObjectMap(item)
		if !ok {
			return nil, fmt.Errorf("inspect designer snapshot warning entry is invalid")
		}
		warning := FormSpecWarning{}
		if code, ok := stringField(warningMap, "code"); ok {
			warning.Code = code
		}
		if message, ok := stringField(warningMap, "message"); ok {
			warning.Message = message
		}
		if control, ok := stringField(warningMap, "control"); ok {
			warning.Control = control
		}
		warnings = append(warnings, warning)
	}
	return warnings, nil
}

func hasObservedFormValues(observed *FormSpecObservedForm) bool {
	return observed != nil && (observed.Caption != nil || observed.Width != nil || observed.Height != nil || observed.InsideWidth != nil || observed.InsideHeight != nil)
}

func hasBuildFormValues(build *FormSpecBuildForm) bool {
	return build != nil && (build.Caption != nil || build.Width != nil || build.Height != nil)
}

func observedFormCaption(observed *FormSpecObservedForm) *string {
	if observed == nil {
		return nil
	}
	return observed.Caption
}

func observedFormWidth(observed *FormSpecObservedForm) *float64 {
	if observed == nil {
		return nil
	}
	return observed.Width
}

func observedFormHeight(observed *FormSpecObservedForm) *float64 {
	if observed == nil {
		return nil
	}
	return observed.Height
}

func hasObservedControlValues(observed *FormSpecObservedControl) bool {
	return observed != nil &&
		(observed.Caption != nil ||
			observed.Text != nil ||
			observed.Value != nil ||
			observed.Left != nil ||
			observed.Top != nil ||
			observed.Width != nil ||
			observed.Height != nil ||
			observed.TabIndex != nil ||
			observed.SelectedIndex != nil ||
			observed.Enabled != nil ||
			observed.Visible != nil ||
			len(observed.List) > 0 ||
			len(observed.Unsupported) > 0 ||
			len(observed.Properties) > 0)
}

func firstStringPtr(values ...*string) *string {
	for _, value := range values {
		if value != nil {
			return value
		}
	}
	return nil
}

func firstFloatPtr(values ...*float64) *float64 {
	for _, value := range values {
		if value != nil {
			return value
		}
	}
	return nil
}

func cloneMap(value map[string]any) map[string]any {
	if len(value) == 0 {
		return nil
	}
	cloned := make(map[string]any, len(value))
	for key, item := range value {
		cloned[key] = item
	}
	return cloned
}

func asObjectMap(value any) (map[string]any, bool) {
	if value == nil {
		return nil, false
	}
	object, ok := value.(map[string]any)
	if !ok {
		return nil, false
	}
	return object, true
}

func asSlice(value any) ([]any, bool) {
	if value == nil {
		return nil, false
	}
	items, ok := value.([]any)
	if !ok {
		return nil, false
	}
	return items, true
}

func stringField(root map[string]any, key string) (string, bool) {
	value, ok := root[key]
	if !ok || value == nil {
		return "", false
	}
	text, ok := value.(string)
	if !ok {
		return "", false
	}
	return text, true
}

func optionalFloatField(root map[string]any, key string) (float64, bool) {
	value, ok := root[key]
	if !ok || value == nil {
		return 0, false
	}
	switch number := value.(type) {
	case float64:
		return number, true
	case float32:
		return float64(number), true
	case int:
		return float64(number), true
	case int32:
		return float64(number), true
	case int64:
		return float64(number), true
	case json.Number:
		parsed, err := number.Float64()
		if err != nil {
			return 0, false
		}
		return parsed, true
	default:
		return 0, false
	}
}

func optionalIntField(root map[string]any, key string) (int, bool) {
	value, ok := root[key]
	if !ok || value == nil {
		return 0, false
	}
	switch number := value.(type) {
	case int:
		return number, true
	case int32:
		return int(number), true
	case int64:
		return int(number), true
	case float64:
		return int(number), true
	case float32:
		return int(number), true
	case json.Number:
		parsed, err := number.Int64()
		if err != nil {
			return 0, false
		}
		return int(parsed), true
	default:
		return 0, false
	}
}

func optionalBoolField(root map[string]any, key string) (bool, bool) {
	value, ok := root[key]
	if !ok || value == nil {
		return false, false
	}
	flag, ok := value.(bool)
	if !ok {
		return false, false
	}
	return flag, true
}

func stringSliceField(root map[string]any, key string) ([]string, bool) {
	items, ok := asSlice(root[key])
	if !ok {
		return nil, false
	}
	values := make([]string, 0, len(items))
	for _, item := range items {
		text, ok := item.(string)
		if !ok {
			return nil, false
		}
		values = append(values, text)
	}
	return values, true
}

func relPath(base, path string) string {
	rel, err := filepath.Rel(base, path)
	if err != nil {
		return path
	}
	return rel
}

func newSpecValidationError(code, message, field string) error {
	return &SpecError{
		Code:    code,
		Message: message,
		Field:   field,
	}
}

func newSpecParseError(input SpecInput, body []byte, err error) error {
	specErr := &SpecError{
		Code:    "spec_parse_failed",
		Message: err.Error(),
		Path:    input.DisplayPath,
		Format:  input.Format,
		Cause:   err,
	}
	if line, column := parseYAMLLineColumn(err.Error()); line > 0 {
		specErr.Line = line
		specErr.Column = column
	} else if line, column := jsonLineColumn(body, err); line > 0 {
		specErr.Line = line
		specErr.Column = column
	}
	specErr.Suggestion = specParseSuggestion(input.Format, body)
	return specErr
}

func parseYAMLLineColumn(message string) (int, int) {
	patterns := []*regexp.Regexp{
		regexp.MustCompile(`line (\d+): column (\d+)`),
		regexp.MustCompile(`line (\d+)`),
	}
	for _, pattern := range patterns {
		matches := pattern.FindStringSubmatch(message)
		if len(matches) == 0 {
			continue
		}
		line, _ := strconv.Atoi(matches[1])
		column := 0
		if len(matches) > 2 {
			column, _ = strconv.Atoi(matches[2])
		}
		return line, column
	}
	return 0, 0
}

func jsonLineColumn(body []byte, err error) (int, int) {
	var syntaxErr *json.SyntaxError
	if !strings.Contains(err.Error(), "invalid") || !asJSONSyntaxError(err, &syntaxErr) {
		return 0, 0
	}
	offset := int(syntaxErr.Offset)
	if offset <= 0 {
		return 0, 0
	}
	line := 1
	column := 1
	for i, b := range body {
		if i >= offset-1 {
			break
		}
		if b == '\n' {
			line++
			column = 1
			continue
		}
		column++
	}
	return line, column
}

func asJSONSyntaxError(err error, target **json.SyntaxError) bool {
	syntaxErr, ok := err.(*json.SyntaxError)
	if !ok {
		return false
	}
	*target = syntaxErr
	return true
}

func specParseSuggestion(format string, body []byte) string {
	if format == "json" {
		return "Fix JSON syntax near the reported location. Check quotes, commas, and trailing delimiters."
	}
	if format != "yaml" {
		return "Fix syntax near the reported location and retry the build."
	}
	text := string(body)
	switch {
	case strings.Contains(text, "caption: -"):
		return `Try quoting scalar strings or use JSON if YAML syntax is uncertain. For an empty caption, use caption: "" rather than caption: -.`
	case strings.Contains(text, ": -"), strings.Contains(text, "\n- "):
		return `Try quoting scalar strings or use JSON if YAML syntax is uncertain. Strings containing ":" or "-" may need quotes.`
	default:
		return `Try quoting scalar strings or use JSON if YAML syntax is uncertain.`
	}
}
