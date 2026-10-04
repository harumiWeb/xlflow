package spec

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"slices"
	"sort"
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
	SchemaVersion      int               `json:"schemaVersion" yaml:"schemaVersion"`
	Kind               string            `json:"kind" yaml:"kind"`
	Basis              string            `json:"basis" yaml:"basis"`
	CoordinateSystem   string            `json:"coordinateSystem,omitempty" yaml:"coordinateSystem,omitempty"`
	Form               FormSpecForm      `json:"form" yaml:"form"`
	Controls           []FormSpecControl `json:"controls" yaml:"controls"`
	Warnings           []FormSpecWarning `json:"warnings" yaml:"warnings"`
	ValidationWarnings []ValidationIssue `json:"-" yaml:"-"`
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
	ClientWidth  *float64 `json:"clientWidth,omitempty" yaml:"clientWidth,omitempty"`
	ClientHeight *float64 `json:"clientHeight,omitempty" yaml:"clientHeight,omitempty"`
}

type FormSpecBuildForm struct {
	Caption      *string  `json:"caption,omitempty" yaml:"caption,omitempty"`
	Width        *float64 `json:"width,omitempty" yaml:"width,omitempty"`
	Height       *float64 `json:"height,omitempty" yaml:"height,omitempty"`
	ClientWidth  *float64 `json:"clientWidth,omitempty" yaml:"clientWidth,omitempty"`
	ClientHeight *float64 `json:"clientHeight,omitempty" yaml:"clientHeight,omitempty"`
}

type FormSpecControl struct {
	ID             string                   `json:"id,omitempty" yaml:"id,omitempty"`
	ParentID       string                   `json:"parentId,omitempty" yaml:"parentId,omitempty"`
	ZIndex         *int                     `json:"zIndex,omitempty" yaml:"zIndex,omitempty"`
	Type           string                   `json:"type" yaml:"type"`
	Name           string                   `json:"name" yaml:"name"`
	ProgID         string                   `json:"progId,omitempty" yaml:"progId,omitempty"`
	Caption        *string                  `json:"caption,omitempty" yaml:"caption,omitempty"`
	Tag            *string                  `json:"tag,omitempty" yaml:"tag,omitempty"`
	ControlTipText *string                  `json:"controlTipText,omitempty" yaml:"controlTipText,omitempty"`
	Accelerator    *string                  `json:"accelerator,omitempty" yaml:"accelerator,omitempty"`
	Text           *string                  `json:"text,omitempty" yaml:"text,omitempty"`
	Value          any                      `json:"value,omitempty" yaml:"value,omitempty"`
	Left           *float64                 `json:"left,omitempty" yaml:"left,omitempty"`
	Top            *float64                 `json:"top,omitempty" yaml:"top,omitempty"`
	Width          *float64                 `json:"width,omitempty" yaml:"width,omitempty"`
	Height         *float64                 `json:"height,omitempty" yaml:"height,omitempty"`
	TabIndex       *int                     `json:"tabIndex,omitempty" yaml:"tabIndex,omitempty"`
	SelectedIndex  *int                     `json:"selectedIndex,omitempty" yaml:"selectedIndex,omitempty"`
	Enabled        *bool                    `json:"enabled,omitempty" yaml:"enabled,omitempty"`
	Visible        *bool                    `json:"visible,omitempty" yaml:"visible,omitempty"`
	List           []string                 `json:"list,omitempty" yaml:"list,omitempty"`
	Tabs           []FormSpecTab            `json:"tabs" yaml:"tabs"`
	Unsupported    []string                 `json:"unsupported,omitempty" yaml:"unsupported,omitempty"`
	Picture        *FormSpecPicture         `json:"picture,omitempty" yaml:"picture,omitempty"`
	Controls       []FormSpecControl        `json:"controls,omitempty" yaml:"controls,omitempty"`
	Properties     map[string]any           `json:"properties,omitempty" yaml:"properties,omitempty"`
	Observed       *FormSpecObservedControl `json:"observed,omitempty" yaml:"observed,omitempty"`
}

// FormSpecPicture describes an Image control picture action. Data contains
// project-root-resolved source bytes for compiler use and is never serialized.
type FormSpecPicture struct {
	Path   string `json:"path,omitempty" yaml:"path,omitempty"`
	Remove bool   `json:"remove,omitzero" yaml:"remove,omitempty"`
	Data   []byte `json:"-" yaml:"-"`
}

func (control FormSpecControl) MarshalYAML() (any, error) {
	type formSpecControlAlias FormSpecControl
	return marshalYAMLWithTabs(formSpecControlAlias(control), control.Tabs)
}

func (control FormSpecControl) MarshalJSON() ([]byte, error) {
	type formSpecControlAlias FormSpecControl
	var tabs json.RawMessage
	var err error
	if control.Tabs != nil {
		tabs, err = json.Marshal(control.Tabs)
		if err != nil {
			return nil, err
		}
	}
	wire := struct {
		formSpecControlAlias
		Tabs json.RawMessage `json:"tabs,omitempty"`
	}{formSpecControlAlias: formSpecControlAlias(control), Tabs: tabs}
	return json.Marshal(wire)
}

// FormSpecTab describes one standalone TabStrip tab.
type FormSpecTab struct {
	Name           string  `json:"name" yaml:"name"`
	Caption        *string `json:"caption,omitempty" yaml:"caption,omitempty"`
	ControlTipText *string `json:"controlTipText,omitempty" yaml:"controlTipText,omitempty"`
	Tag            *string `json:"tag,omitempty" yaml:"tag,omitempty"`
	Accelerator    *string `json:"accelerator,omitempty" yaml:"accelerator,omitempty"`
	Enabled        *bool   `json:"enabled,omitempty" yaml:"enabled,omitempty"`
	Visible        *bool   `json:"visible,omitempty" yaml:"visible,omitempty"`
}

type FormSpecObservedControl struct {
	Caption        *string        `json:"caption,omitempty" yaml:"caption,omitempty"`
	Tag            *string        `json:"tag,omitempty" yaml:"tag,omitempty"`
	ControlTipText *string        `json:"controlTipText,omitempty" yaml:"controlTipText,omitempty"`
	Accelerator    *string        `json:"accelerator,omitempty" yaml:"accelerator,omitempty"`
	Text           *string        `json:"text,omitempty" yaml:"text,omitempty"`
	Value          any            `json:"value,omitempty" yaml:"value,omitempty"`
	Left           *float64       `json:"left,omitempty" yaml:"left,omitempty"`
	Top            *float64       `json:"top,omitempty" yaml:"top,omitempty"`
	Width          *float64       `json:"width,omitempty" yaml:"width,omitempty"`
	Height         *float64       `json:"height,omitempty" yaml:"height,omitempty"`
	TabIndex       *int           `json:"tabIndex,omitempty" yaml:"tabIndex,omitempty"`
	SelectedIndex  *int           `json:"selectedIndex,omitempty" yaml:"selectedIndex,omitempty"`
	Enabled        *bool          `json:"enabled,omitempty" yaml:"enabled,omitempty"`
	Visible        *bool          `json:"visible,omitempty" yaml:"visible,omitempty"`
	List           []string       `json:"list,omitempty" yaml:"list,omitempty"`
	Tabs           []FormSpecTab  `json:"tabs" yaml:"tabs"`
	Unsupported    []string       `json:"unsupported,omitempty" yaml:"unsupported,omitempty"`
	Properties     map[string]any `json:"properties,omitempty" yaml:"properties,omitempty"`
}

func (control FormSpecObservedControl) MarshalYAML() (any, error) {
	type formSpecObservedControlAlias FormSpecObservedControl
	return marshalYAMLWithTabs(formSpecObservedControlAlias(control), control.Tabs)
}

func (control FormSpecObservedControl) MarshalJSON() ([]byte, error) {
	type formSpecObservedControlAlias FormSpecObservedControl
	var tabs json.RawMessage
	var err error
	if control.Tabs != nil {
		tabs, err = json.Marshal(control.Tabs)
		if err != nil {
			return nil, err
		}
	}
	wire := struct {
		formSpecObservedControlAlias
		Tabs json.RawMessage `json:"tabs,omitempty"`
	}{formSpecObservedControlAlias: formSpecObservedControlAlias(control), Tabs: tabs}
	return json.Marshal(wire)
}

func marshalYAMLWithTabs(value any, tabs []FormSpecTab) (any, error) {
	body, err := yaml.Marshal(value)
	if err != nil {
		return nil, err
	}
	var document yaml.Node
	if err := yaml.Unmarshal(body, &document); err != nil {
		return nil, err
	}
	if len(document.Content) != 1 {
		return nil, errors.New("FormSpec YAML value did not produce one document node")
	}
	root := document.Content[0]
	for index := 0; index+1 < len(root.Content); index += 2 {
		if root.Content[index].Value != "tabs" {
			continue
		}
		if tabs == nil {
			root.Content = append(root.Content[:index], root.Content[index+2:]...)
		}
		return root, nil
	}
	return nil, errors.New("FormSpec YAML value is missing its tabs field")
}

type FormSpecWarning struct {
	Code    string `json:"code,omitempty" yaml:"code,omitempty"`
	Message string `json:"message,omitempty" yaml:"message,omitempty"`
	Control string `json:"control,omitempty" yaml:"control,omitempty"`
}

const (
	UnsupportedControlPlaceholderType              = "Control"
	UnsupportedControlTypeProperty                 = "controlType"
	CompatibilityArtifactUnsynchronizedWarningCode = "compatibility_artifact_unsynchronized"
	formDimensionValidationCode                    = "UFV016"
	formBuildDimensionConflictCode                 = "UFV017"
	excelClientDimensionValidationCode             = "UFV018"
	tabNameValidationCode                          = "UFV019"
	selectedIndexValidationCode                    = "UFV020"
)

type SpecError struct {
	Code       string
	Message    string
	Path       string
	Format     string
	Line       int
	Column     int
	Field      string
	Suggestion string
	Issues     []ValidationIssue
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

type Severity string

const (
	SeverityError   Severity = "error"
	SeverityWarning Severity = "warning"
)

type ValidationIssue struct {
	Code       string       `json:"code,omitempty"`
	Severity   Severity     `json:"severity,omitempty"`
	Message    string       `json:"message,omitempty"`
	Field      string       `json:"field,omitempty"`
	Suggestion string       `json:"suggestion,omitempty"`
	Support    SupportLevel `json:"support,omitempty"`
}

func ResolveSnapshotOutput(root, outPath string) (SnapshotOutput, error) {
	trimmed := strings.TrimSpace(outPath)
	if trimmed == "" {
		return SnapshotOutput{}, fmt.Errorf("--out is required")
	}
	resolved := normalizeCLIPath(trimmed)
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
	body, err := MarshalSnapshot(output.Format, spec)
	if err != nil {
		return err
	}
	return os.WriteFile(output.Path, body, 0o644)
}

// MarshalSnapshot serializes a canonical UserForm snapshot without touching
// the filesystem. Transactional callers use it to finish validation and
// serialization before publishing any source artifact.
func MarshalSnapshot(format string, spec FormSpec) ([]byte, error) {
	var body []byte
	var err error
	switch format {
	case "json":
		body, err = json.MarshalIndent(spec, "", "  ")
		if err == nil {
			body = append(body, '\n')
		}
	case "yaml":
		body, err = yaml.Marshal(spec)
	default:
		err = fmt.Errorf("unsupported snapshot format %q", format)
	}
	if err != nil {
		return nil, err
	}
	return body, nil
}

func ResolveSpecInput(root, specPath string) (SpecInput, error) {
	trimmed := strings.TrimSpace(specPath)
	if trimmed == "" {
		return SpecInput{}, fmt.Errorf("spec path is required")
	}
	resolved := normalizeCLIPath(trimmed)
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

func normalizeCLIPath(path string) string {
	return strings.ReplaceAll(path, `\`, string(filepath.Separator))
}

func LoadFormSpec(input SpecInput) (FormSpec, error) {
	body, err := os.ReadFile(input.Path)
	if err != nil {
		return FormSpec{}, err
	}
	issues, err := ValidateFormSpecSource(input, body)
	if err != nil {
		return FormSpec{}, err
	}
	if hasValidationErrors(issues) {
		return FormSpec{}, newSpecValidationIssuesError(input, issues)
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
	structIssues := ValidateFormSpecStrict(spec)
	if hasValidationErrors(structIssues) {
		err := newSpecValidationIssuesError(input, structIssues)
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
	spec.ValidationWarnings = validationWarnings(append(issues, structIssues...))
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
	issues := ValidateFormSpecStrict(spec)
	if hasValidationErrors(issues) {
		return newSpecValidationErrorFromIssue(firstValidationError(issues), issues)
	}
	return nil
}

// ValidateFormSpecForAuthoring rejects snapshot-only control placeholders and
// client dimensions that cannot be authored safely by an Excel writer.
func ValidateFormSpecForAuthoring(input SpecInput, spec FormSpec) error {
	issues := make([]ValidationIssue, 0)
	for i, control := range spec.Controls {
		if !IsUnsupportedControlPlaceholder(control) {
			continue
		}
		path := fmt.Sprintf("controls[%d]", i)
		issues = append(issues, validationIssue(
			"UFV006",
			SeverityError,
			fmt.Sprintf("%s.type %q is a snapshot-only placeholder and cannot be authored.", path, control.Type),
			path+".type",
			"Replace the placeholder with a supported built-in type or provide the control's real custom progId.",
			"",
		))
	}
	if build := spec.Form.Build; build != nil {
		for _, field := range []struct {
			path  string
			value *float64
		}{
			{path: "form.build.clientWidth", value: build.ClientWidth},
			{path: "form.build.clientHeight", value: build.ClientHeight},
		} {
			if field.value == nil {
				continue
			}
			issues = append(issues, validationIssue(
				excelClientDimensionValidationCode,
				SeverityError,
				fmt.Sprintf("%s is a pure-Go build client dimension and is not supported by Excel-backed form build; refusing it instead of ignoring it.", field.path),
				field.path,
				"Remove the client dimension before using `form build` or use the pure-Go generator.",
				SupportLevelSupported,
			))
		}
	}
	if hasValidationErrors(issues) {
		return newSpecValidationIssuesError(input, issues)
	}
	return nil
}

func ValidateFormSpecSource(input SpecInput, body []byte) ([]ValidationIssue, error) {
	value, err := decodeSpecSource(input, body)
	if err != nil {
		return nil, err
	}
	root, ok := asObjectMap(value)
	if !ok {
		return []ValidationIssue{validationIssue("UFV002", SeverityError, "UserForm spec root must be an object.", "", "", "")}, nil
	}
	return validateRawFormSpec(root), nil
}

func ValidateFormSpecStrict(spec FormSpec) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	if spec.SchemaVersion != 1 {
		issues = append(issues, invalidFixedValueIssue("schemaVersion", "1"))
	}
	if spec.Kind != "xlflow.userform" {
		issues = append(issues, invalidFixedValueIssue("kind", `"xlflow.userform"`))
	}
	if strings.TrimSpace(spec.Basis) != "designer" {
		issues = append(issues, invalidFixedValueIssue("basis", `"designer"`))
	}
	if strings.TrimSpace(spec.Form.Name) == "" {
		issues = append(issues, requiredFieldIssue("form.name"))
	}
	issues = append(issues, validateFormDimensions(spec.Form)...)
	ids := make(map[string]struct{}, len(spec.Controls))
	controlsByID := make(map[string]FormSpecControl, len(spec.Controls))
	for i, control := range spec.Controls {
		path := fmt.Sprintf("controls[%d]", i)
		issues = append(issues, ValidateFormSpecControlIssues(control, path)...)
		if strings.TrimSpace(control.ID) == "" {
			continue
		}
		if _, exists := ids[control.ID]; exists {
			issues = append(issues, validationIssue("UFV007", SeverityError, fmt.Sprintf("%s.id %q is duplicated.", path, control.ID), path+".id", "Use a unique stable id for each control.", ""))
			continue
		}
		ids[control.ID] = struct{}{}
		controlsByID[control.ID] = control
	}
	parentByID := make(map[string]string, len(spec.Controls))
	parentFieldByID := make(map[string]string, len(spec.Controls))
	pageCountByParentID := make(map[string]int, len(spec.Controls))
	pageTopologyKnown := make(map[string]bool, len(spec.Controls))
	for i, control := range spec.Controls {
		if strings.EqualFold(strings.TrimSpace(control.Type), "Page") && strings.TrimSpace(control.ParentID) != "" {
			pageCountByParentID[control.ParentID]++
			pageTopologyKnown[control.ParentID] = true
		}
		if strings.EqualFold(strings.TrimSpace(control.Type), "MultiPage") && control.Controls != nil {
			pageTopologyKnown[control.ID] = true
			for _, child := range control.Controls {
				if strings.EqualFold(strings.TrimSpace(child.Type), "Page") {
					pageCountByParentID[control.ID]++
				}
			}
		}
		if strings.TrimSpace(control.ParentID) == "" {
			if strings.EqualFold(strings.TrimSpace(control.Type), "Page") {
				issues = append(issues, requiredPageParentIssue(fmt.Sprintf("controls[%d].parentId", i)))
			}
			continue
		}
		field := fmt.Sprintf("controls[%d].parentId", i)
		if control.ParentID == control.ID {
			issues = append(issues, validationIssue("UFV009", SeverityError, fmt.Sprintf("%s must not reference the same control.", field), field, "Remove parentId or point it at a container control.", ""))
			continue
		}
		parent, ok := controlsByID[control.ParentID]
		if !ok {
			issues = append(issues, validationIssue("UFV008", SeverityError, fmt.Sprintf("%s %q was not found.", field, control.ParentID), field, "Use the id of an existing container control.", ""))
			continue
		}
		if allowed, knownParentControl := FormSpecControlParentAllowsChild(parent, control); knownParentControl && !allowed {
			issues = append(issues, controlParentValidationIssue(field, parent, control))
		}
		parentByID[control.ID] = control.ParentID
		parentFieldByID[control.ID] = field
	}
	for i, control := range spec.Controls {
		if !strings.EqualFold(strings.TrimSpace(control.Type), "MultiPage") || !pageTopologyKnown[control.ID] {
			continue
		}
		count := pageCountByParentID[control.ID]
		tabs := make([]FormSpecTab, count)
		path := fmt.Sprintf("controls[%d]", i)
		issues = append(issues, validateSelectedIndex(control.SelectedIndex, tabs, path+".selectedIndex", true)...)
		if control.Observed != nil {
			issues = append(issues, validateSelectedIndex(control.Observed.SelectedIndex, tabs, path+".observed.selectedIndex", true)...)
		}
	}
	issues = append(issues, parentCycleIssues(parentByID, parentFieldByID)...)
	return issues
}

func ValidateFormSpecControl(control FormSpecControl, path string) error {
	issues := ValidateFormSpecControlIssues(control, path)
	if hasValidationErrors(issues) {
		return newSpecValidationErrorFromIssue(firstValidationError(issues), issues)
	}
	return nil
}

func ValidateFormSpecControlIssues(control FormSpecControl, path string) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	if strings.TrimSpace(control.ID) == "" {
		issues = append(issues, requiredFieldIssue(path+".id"))
	}
	if strings.TrimSpace(control.Name) == "" {
		issues = append(issues, requiredFieldIssue(path+".name"))
	}
	if strings.TrimSpace(control.Type) == "" {
		issues = append(issues, requiredFieldIssue(path+".type"))
	}
	if IsUnsupportedControlPlaceholder(control) {
		issues = append(issues, unsupportedControlPlaceholderIssue(path))
	} else if _, err := ControlProgID(control); err != nil {
		issues = append(issues, validationIssue("UFV006", SeverityError, fmt.Sprintf("%s: %v.", path, err), path+".type", "Use a supported built-in type or provide a custom progId.", ""))
	}
	if strings.TrimSpace(control.Type) != "" && strings.TrimSpace(control.ProgID) != "" {
		if progControl, ok := LookupControlContractByProgID(control.ProgID); ok && !strings.EqualFold(strings.TrimSpace(control.Type), progControl.Type) {
			issues = append(issues, validationIssue("UFV012", SeverityError, fmt.Sprintf("%s.progId %q is for %s, not %s.", path, control.ProgID, progControl.Type, control.Type), path+".progId", "Use the ProgID that matches type or change type to match the ProgID.", ""))
		}
	}
	if control.Picture != nil {
		issues = append(issues, validatePictureAction(control.Type, control.Picture, path+".picture")...)
	}
	if strings.EqualFold(strings.TrimSpace(control.Type), "Page") {
		for _, field := range []struct {
			name  string
			value *float64
		}{
			{name: "left", value: control.Left},
			{name: "top", value: control.Top},
			{name: "width", value: control.Width},
			{name: "height", value: control.Height},
		} {
			if field.value != nil {
				issues = append(issues, invalidControlPropertyIssue(path+"."+field.name, control.Type))
			}
		}
	}
	if control.Tabs != nil {
		if !strings.EqualFold(strings.TrimSpace(control.Type), "TabStrip") {
			issues = append(issues, invalidControlPropertyIssue(path+".tabs", control.Type))
		} else {
			issues = append(issues, validateFormSpecTabs(control.Tabs, path+".tabs")...)
		}
	}
	if isPageSelectionControl(control.Type) {
		issues = append(issues, validateSelectedIndex(control.SelectedIndex, control.Tabs, path+".selectedIndex", strings.EqualFold(strings.TrimSpace(control.Type), "TabStrip") && control.Tabs != nil)...)
		if control.Observed != nil {
			if control.Observed.Tabs != nil && !strings.EqualFold(strings.TrimSpace(control.Type), "TabStrip") {
				issues = append(issues, invalidControlPropertyIssue(path+".observed.tabs", control.Type))
			} else {
				issues = append(issues, validateFormSpecTabs(control.Observed.Tabs, path+".observed.tabs")...)
			}
			issues = append(issues, validateSelectedIndex(control.Observed.SelectedIndex, control.Observed.Tabs, path+".observed.selectedIndex", strings.EqualFold(strings.TrimSpace(control.Type), "TabStrip") && control.Observed.Tabs != nil)...)
		}
	} else if control.Observed != nil && control.Observed.Tabs != nil {
		issues = append(issues, invalidControlPropertyIssue(path+".observed.tabs", control.Type))
	}
	return issues
}

func validateFormSpecTabs(tabs []FormSpecTab, path string) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	seen := make([]string, 0, len(tabs))
	for i, tab := range tabs {
		field := fmt.Sprintf("%s[%d].name", path, i)
		name := strings.TrimSpace(tab.Name)
		if name == "" {
			issues = append(issues, requiredFieldIssue(field))
			continue
		}
		if first := matchingTabName(seen, name); first != "" {
			issues = append(issues, validationIssue(tabNameValidationCode, SeverityError, fmt.Sprintf("%s duplicates tab name %q from %s.", field, name, first), field, "Use a tab name that is unique within this TabStrip, ignoring case.", ""))
			continue
		}
		seen = append(seen, name)
	}
	return issues
}

func validateSelectedIndex(index *int, tabs []FormSpecTab, path string, bounded bool) []ValidationIssue {
	if index == nil {
		return nil
	}
	if *index < -1 || bounded && ((*index == -1 && len(tabs) > 0) || *index >= len(tabs)) {
		return []ValidationIssue{validationIssue(selectedIndexValidationCode, SeverityError, fmt.Sprintf("%s must be -1 or a valid zero-based index.", path), path, "Use -1 for no selection or an index within the declared tabs.", "")}
	}
	return nil
}

func isPageSelectionControl(controlType string) bool {
	return strings.EqualFold(strings.TrimSpace(controlType), "MultiPage") || strings.EqualFold(strings.TrimSpace(controlType), "TabStrip")
}

func invalidControlPropertyIssue(path, controlType string) ValidationIssue {
	return validationIssue("UFV005", SeverityError, fmt.Sprintf("%s is not valid for control type %s.", path, controlType), path, "Remove the property or use the control type that supports it.", "")
}

func requiredPageParentIssue(path string) ValidationIssue {
	return validationIssue("UFV011", SeverityError, fmt.Sprintf("%s is required for Page controls; a Page must belong to a MultiPage.", path), path, "Set parentId to the owning MultiPage control.", "")
}

func controlParentValidationIssue(path string, parent, child FormSpecControl) ValidationIssue {
	if canContainChildren, known := FormSpecControlCanContainChildren(parent); known && !canContainChildren {
		return validationIssue("UFV011", SeverityError, fmt.Sprintf("%s %q references non-container control %q.", path, parent.ID, parent.Name), path, "Use a control allowed to contain this child as the parent.", "")
	}
	return validationIssue("UFV011", SeverityError, fmt.Sprintf("%s parent %q is not allowed for control type %s.", path, parent.Name, child.Type), path, "MultiPage controls may contain only Page controls, and Page controls must belong to a MultiPage.", "")
}

// IsUnsupportedControlPlaceholder reports whether a control is a loss-aware
// snapshot placeholder rather than an authorable control definition.
func IsUnsupportedControlPlaceholder(control FormSpecControl) bool {
	return strings.EqualFold(strings.TrimSpace(control.Type), UnsupportedControlPlaceholderType) &&
		strings.TrimSpace(control.ProgID) == "" &&
		slices.Contains(control.Unsupported, UnsupportedControlTypeProperty)
}

func unsupportedControlPlaceholderIssue(path string) ValidationIssue {
	return validationIssue(
		"UFV015",
		SeverityWarning,
		fmt.Sprintf("%s is a snapshot-only placeholder for a control whose type could not be recovered.", path),
		path+".type",
		"Preserve it for review, or replace it with a supported type and real progId before building.",
		SupportLevelSnapshotOnly,
	)
}

func ControlProgID(control FormSpecControl) (string, error) {
	if progID := strings.TrimSpace(control.ProgID); progID != "" {
		return progID, nil
	}
	progID, ok := BuiltInControlProgID(control.Type)
	if !ok {
		return "", fmt.Errorf("unsupported control type %q", control.Type)
	}
	return progID, nil
}

func decodeSpecSource(input SpecInput, body []byte) (any, error) {
	switch input.Format {
	case "json":
		decoder := json.NewDecoder(bytes.NewReader(body))
		decoder.UseNumber()
		var value any
		if err := decoder.Decode(&value); err != nil {
			return nil, newSpecParseError(input, body, err)
		}
		var trailing any
		if err := decoder.Decode(&trailing); err != io.EOF {
			return nil, newSpecParseError(input, body, fmt.Errorf("invalid JSON document"))
		}
		return normalizeDecodedValue(value), nil
	case "yaml":
		var node yaml.Node
		if err := yaml.Unmarshal(body, &node); err != nil {
			return nil, newSpecParseError(input, body, err)
		}
		var value any
		if err := node.Decode(&value); err != nil {
			return nil, newSpecParseError(input, body, err)
		}
		return normalizeDecodedValue(value), nil
	default:
		return nil, fmt.Errorf("unsupported form spec format %q", input.Format)
	}
}

func validateRawFormSpec(root map[string]any) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	contract := UserFormContract()
	issues = append(issues, validateObjectProperties(root, contract.DocumentProperties, "", nil)...)
	form, ok := root["form"]
	if ok {
		if formMap, formOK := asObjectMap(form); formOK {
			issues = append(issues, validateObjectProperties(formMap, contract.FormProperties, "form", nil)...)
			issues = append(issues, validateRawFormSubobject(formMap, "build", formBuildProperties(), "form.build")...)
			issues = append(issues, validateRawFormSubobject(formMap, "observed", formObservedProperties(), "form.observed")...)
			issues = append(issues, validateRawFormDimensions(formMap)...)
		}
	}
	rawControls, controlsOK := root["controls"]
	if controlsOK {
		flatControls, controlIssues := validateRawControls(rawControls, "controls", "")
		issues = append(issues, controlIssues...)
		issues = append(issues, validateRawControlStructure(flatControls)...)
	}
	return issues
}

func validateRawFormSubobject(root map[string]any, key string, properties map[string]PropertyContract, path string) []ValidationIssue {
	value, ok := root[key]
	if !ok || value == nil {
		return nil
	}
	object, ok := asObjectMap(value)
	if !ok {
		return nil
	}
	return validateObjectProperties(object, properties, path, nil)
}

type rawControlRef struct {
	Path                  string
	ID                    string
	ParentID              string
	Name                  string
	Type                  string
	ProgID                string
	SelectedIndex         *float64
	ObservedSelectedIndex *float64
	TabsKnown             bool
	TabCount              int
	ChildrenKnown         bool
}

func validateRawControls(value any, path, inheritedParentID string) ([]rawControlRef, []ValidationIssue) {
	items, ok := asSlice(value)
	if !ok {
		return nil, nil
	}
	refs := make([]rawControlRef, 0, len(items))
	issues := make([]ValidationIssue, 0)
	for i, item := range items {
		itemPath := fmt.Sprintf("%s[%d]", path, i)
		controlMap, ok := asObjectMap(item)
		if !ok {
			issues = append(issues, validationIssue("UFV002", SeverityError, fmt.Sprintf("%s must be an object.", itemPath), itemPath, "", ""))
			continue
		}
		refsForControl, issuesForControl := validateRawControl(controlMap, itemPath, inheritedParentID)
		refs = append(refs, refsForControl...)
		issues = append(issues, issuesForControl...)
	}
	return refs, issues
}

func validateRawControl(controlMap map[string]any, path, inheritedParentID string) ([]rawControlRef, []ValidationIssue) {
	issues := make([]ValidationIssue, 0)
	controlType, _ := stringField(controlMap, "type")
	progID, _ := stringField(controlMap, "progId")
	issues = append(issues, validateRawControlProperties(controlMap, controlType, progID, path)...)
	id, _ := stringField(controlMap, "id")
	name, _ := stringField(controlMap, "name")
	parentID, _ := stringField(controlMap, "parentId")
	if strings.TrimSpace(parentID) == "" {
		parentID = inheritedParentID
	}
	ref := rawControlRef{
		Path:     path,
		ID:       strings.TrimSpace(id),
		ParentID: strings.TrimSpace(parentID),
		Name:     strings.TrimSpace(name),
		Type:     strings.TrimSpace(controlType),
		ProgID:   strings.TrimSpace(progID),
	}
	if value, ok := lookupRawField(controlMap, "selectedIndex"); ok {
		ref.SelectedIndex = rawSelectedIndex(value)
	}
	if rawTabs, ok := lookupRawField(controlMap, "tabs"); ok {
		if tabs, valid := asSlice(rawTabs); valid {
			ref.TabsKnown = true
			ref.TabCount = len(tabs)
		}
	}
	refs := []rawControlRef{ref}
	if observed, ok := asObjectMap(controlMap["observed"]); ok {
		issues = append(issues, validateRawObservedControlProperties(observed, ref.Type, ref.ProgID, path+".observed")...)
		refs[0].ObservedSelectedIndex = rawSelectedIndex(observed["selectedIndex"])
	}
	if children, ok := controlMap["controls"]; ok {
		refs[0].ChildrenKnown = true
		issues = append(issues, validationIssue("UFV013", SeverityWarning, fmt.Sprintf("%s.controls is a legacy nested control structure.", path), path+".controls", "Prefer the canonical flat controls array with parentId references.", SupportLevelSnapshotOnly))
		childRefs, childIssues := validateRawControls(children, path+".controls", ref.ID)
		refs = append(refs, childRefs...)
		issues = append(issues, childIssues...)
	}
	return refs, issues
}

func validateRawControlProperties(controlMap map[string]any, controlType, progID, path string) []ValidationIssue {
	contract := UserFormContract()
	allowed := clonePropertyMap(contract.CommonControlProperties)
	_, builtInType := LookupControlContract(controlType)
	if builtInType {
		allowed = ControlProperties(controlType)
	} else if isCustomProgID(progID) {
		addGenericCustomControlProperties(allowed, SupportLevelCustomUnchecked)
	}
	issues := validateObjectProperties(controlMap, allowed, path, nil)
	if value, ok := lookupRawField(controlMap, "picture"); ok {
		issues = append(issues, validateRawPicture(value, controlType, path+".picture")...)
	}
	if rawTabs, ok := lookupRawField(controlMap, "tabs"); ok {
		issues = append(issues, validateRawFormSpecTabs(rawTabs, path+".tabs")...)
	}
	issues = append(issues, validateRawSelectedIndex(controlMap, controlType, path)...)
	if builtInType {
		markUnsupportedControlProperties(issues, controlType)
	}
	if strings.TrimSpace(controlType) != "" && strings.TrimSpace(progID) != "" {
		if progControl, ok := LookupControlContractByProgID(progID); ok && !strings.EqualFold(strings.TrimSpace(controlType), progControl.Type) {
			issues = append(issues, validationIssue("UFV012", SeverityError, fmt.Sprintf("%s.progId %q is for %s, not %s.", path, progID, progControl.Type, controlType), path+".progId", "Use the ProgID that matches type or change type to match the ProgID.", ""))
		}
	}
	if strings.TrimSpace(controlType) != "" && !builtInType {
		if isRawUnsupportedControlPlaceholder(controlMap, controlType, progID) {
			issues = append(issues, unsupportedControlPlaceholderIssue(path))
		} else if strings.TrimSpace(progID) == "" {
			issues = append(issues, validationIssue("UFV006", SeverityError, fmt.Sprintf("%s.type %q is not a supported built-in control type.", path, controlType), path+".type", "Use a supported built-in type or provide a custom progId.", ""))
		} else if _, knownProgID := LookupControlContractByProgID(progID); !knownProgID {
			issues = append(issues, validationIssue("UFV014", SeverityWarning, fmt.Sprintf("%s uses custom control type %q with unchecked ProgID %q.", path, controlType, progID), path+".progId", "Only common structural fields and the properties bag are validated for custom controls.", SupportLevelCustomUnchecked))
		}
	}
	return issues
}

func validateRawPicture(value any, controlType, path string) []ValidationIssue {
	if !strings.EqualFold(strings.TrimSpace(controlType), "Image") {
		return []ValidationIssue{invalidControlPropertyIssue(path, controlType)}
	}
	object, ok := asObjectMap(value)
	if !ok {
		return []ValidationIssue{validationIssue("UFV005", SeverityError, fmt.Sprintf("%s must be an object.", path), path, "Use {path: relative-file} or {remove: true}.", "")}
	}
	fields := make(map[string]any, len(object))
	issues := make([]ValidationIssue, 0)
	for key, item := range object {
		name := strings.ToLower(strings.TrimSpace(key))
		if name != "path" && name != "remove" {
			issues = append(issues, validationIssue("UFV001", SeverityError, fmt.Sprintf("%s.%s is not defined by the picture action.", path, key), path+"."+key, "Use only path or remove.", ""))
			continue
		}
		if _, duplicate := fields[name]; duplicate {
			issues = append(issues, validationIssue("UFV001", SeverityError, fmt.Sprintf("%s.%s is duplicated ignoring case.", path, key), path+"."+key, "Keep one spelling for each picture action field.", ""))
			continue
		}
		fields[name] = item
	}
	pathValue, hasPath := fields["path"]
	removeValue, hasRemove := fields["remove"]
	switch {
	case hasPath && hasRemove:
		issues = append(issues, validationIssue("UFV005", SeverityError, fmt.Sprintf("%s must contain exactly one of path or remove.", path), path, "Use {path: relative-file} or {remove: true}, not both.", ""))
	case hasPath:
		text, ok := pathValue.(string)
		if !ok || strings.TrimSpace(text) == "" {
			issues = append(issues, validationIssue("UFV005", SeverityError, fmt.Sprintf("%s.path must be a non-empty string.", path), path+".path", "Provide a project-root-relative BMP or JPEG path.", ""))
		}
	case hasRemove:
		remove, ok := removeValue.(bool)
		if !ok || !remove {
			issues = append(issues, validationIssue("UFV005", SeverityError, fmt.Sprintf("%s.remove must be true.", path), path+".remove", "Use remove: true to clear the picture.", ""))
		}
	default:
		issues = append(issues, validationIssue("UFV005", SeverityError, fmt.Sprintf("%s must contain path or remove: true.", path), path, "Use {path: relative-file} or {remove: true}.", ""))
	}
	return issues
}

func validatePictureAction(controlType string, picture *FormSpecPicture, path string) []ValidationIssue {
	if !strings.EqualFold(strings.TrimSpace(controlType), "Image") {
		return []ValidationIssue{invalidControlPropertyIssue(path, controlType)}
	}
	if picture.Remove {
		if strings.TrimSpace(picture.Path) != "" {
			return []ValidationIssue{validationIssue("UFV005", SeverityError, fmt.Sprintf("%s cannot combine path and remove.", path), path, "Use exactly one picture action.", "")}
		}
		return nil
	}
	if strings.TrimSpace(picture.Path) == "" {
		return []ValidationIssue{validationIssue("UFV005", SeverityError, fmt.Sprintf("%s.path must be a non-empty string.", path), path+".path", "Provide a project-root-relative BMP or JPEG path.", "")}
	}
	return nil
}

func isRawUnsupportedControlPlaceholder(controlMap map[string]any, controlType, progID string) bool {
	if !strings.EqualFold(strings.TrimSpace(controlType), UnsupportedControlPlaceholderType) || strings.TrimSpace(progID) != "" {
		return false
	}
	unsupported, ok := stringSliceField(controlMap, "unsupported")
	return ok && slices.Contains(unsupported, UnsupportedControlTypeProperty)
}

func validateRawObservedControlProperties(controlMap map[string]any, controlType, progID, path string) []ValidationIssue {
	contract := UserFormContract()
	allowed := map[string]PropertyContract{}
	for _, key := range []string{"left", "top", "width", "height", "tabIndex", "enabled", "visible", "unsupported", "properties"} {
		if property, ok := lookupProperty(contract.CommonControlProperties, key); ok {
			allowed[key] = property
		}
	}
	if builtInControl, ok := LookupControlContract(controlType); ok {
		for key, property := range builtInControl.Properties {
			allowed[key] = property
		}
	} else if isCustomProgID(progID) {
		addGenericCustomControlProperties(allowed, SupportLevelCustomUnchecked)
	}
	issues := validateObjectProperties(controlMap, allowed, path, nil)
	if rawTabs, ok := lookupRawField(controlMap, "tabs"); ok {
		issues = append(issues, validateRawFormSpecTabs(rawTabs, path+".tabs")...)
	}
	issues = append(issues, validateRawSelectedIndex(controlMap, controlType, path)...)
	if _, ok := LookupControlContract(controlType); ok {
		markUnsupportedControlProperties(issues, controlType)
	}
	return issues
}

func validateRawControlStructure(controls []rawControlRef) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	ids := make(map[string]rawControlRef, len(controls))
	parentByID := make(map[string]string, len(controls))
	parentFieldByID := make(map[string]string, len(controls))
	for _, control := range controls {
		if control.ID == "" {
			continue
		}
		if existing, exists := ids[control.ID]; exists {
			issues = append(issues, validationIssue("UFV007", SeverityError, fmt.Sprintf("%s.id %q is duplicated; first seen at %s.id.", control.Path, control.ID, existing.Path), control.Path+".id", "Use a unique stable id for each control.", ""))
			continue
		}
		ids[control.ID] = control
	}
	for _, control := range controls {
		if control.ParentID == "" {
			if strings.EqualFold(control.Type, "Page") {
				issues = append(issues, requiredPageParentIssue(control.Path+".parentId"))
			}
			continue
		}
		field := control.Path + ".parentId"
		if control.ID != "" && control.ParentID == control.ID {
			issues = append(issues, validationIssue("UFV009", SeverityError, fmt.Sprintf("%s must not reference the same control.", field), field, "Remove parentId or point it at a container control.", ""))
			continue
		}
		parent, ok := ids[control.ParentID]
		if !ok {
			issues = append(issues, validationIssue("UFV008", SeverityError, fmt.Sprintf("%s %q was not found.", field, control.ParentID), field, "Use the id of an existing container control.", ""))
			continue
		}
		parentSpec := FormSpecControl{ID: parent.ID, Name: parent.Name, Type: parent.Type, ProgID: parent.ProgID}
		childSpec := FormSpecControl{ID: control.ID, Name: control.Name, Type: control.Type, ProgID: control.ProgID}
		if allowed, known := FormSpecControlParentAllowsChild(parentSpec, childSpec); known && !allowed {
			issues = append(issues, controlParentValidationIssue(field, parentSpec, childSpec))
		}
		if control.ID != "" {
			parentByID[control.ID] = control.ParentID
			parentFieldByID[control.ID] = field
		}
	}
	pageCountByParentID := make(map[string]int, len(controls))
	pageTopologyKnown := make(map[string]bool, len(controls))
	for _, control := range controls {
		if strings.EqualFold(control.Type, "MultiPage") && control.ChildrenKnown {
			pageTopologyKnown[control.ID] = true
		}
		if strings.EqualFold(control.Type, "Page") && control.ParentID != "" {
			pageCountByParentID[control.ParentID]++
			pageTopologyKnown[control.ParentID] = true
		}
	}
	for _, control := range controls {
		if !strings.EqualFold(control.Type, "MultiPage") || !pageTopologyKnown[control.ID] {
			continue
		}
		count := pageCountByParentID[control.ID]
		for _, item := range []struct {
			path  string
			index *float64
		}{{control.Path + ".selectedIndex", control.SelectedIndex}, {control.Path + ".observed.selectedIndex", control.ObservedSelectedIndex}} {
			if item.index == nil {
				continue
			}
			invalid := count == 0 && *item.index != -1 || count > 0 && (*item.index < 0 || *item.index >= float64(count))
			if invalid {
				issues = append(issues, selectedIndexIssue(item.path))
			}
		}
	}
	return append(issues, parentCycleIssues(parentByID, parentFieldByID)...)
}

func validateRawFormSpecTabs(value any, path string) []ValidationIssue {
	items, ok := asSlice(value)
	if !ok {
		return nil
	}
	properties := UserFormContract().TabProperties
	issues := make([]ValidationIssue, 0)
	seen := make([]string, 0, len(items))
	for i, item := range items {
		itemPath := fmt.Sprintf("%s[%d]", path, i)
		object, ok := asObjectMap(item)
		if !ok {
			continue // The enclosing object-array check reports the malformed entry.
		}
		issues = append(issues, validateObjectProperties(object, properties, itemPath, nil)...)
		name, _ := stringField(object, "name")
		name = strings.TrimSpace(name)
		if name == "" {
			continue // Required-field validation above reports the missing name.
		}
		if first := matchingTabName(seen, name); first != "" {
			field := itemPath + ".name"
			issues = append(issues, validationIssue(tabNameValidationCode, SeverityError, fmt.Sprintf("%s duplicates tab name %q from an earlier tab in this TabStrip.", field, first), field, "Use a tab name that is unique within this TabStrip, ignoring case.", ""))
			continue
		}
		seen = append(seen, name)
	}
	return issues
}

func matchingTabName(names []string, name string) string {
	for _, candidate := range names {
		if strings.EqualFold(candidate, name) {
			return candidate
		}
	}
	return ""
}

func validateRawSelectedIndex(values map[string]any, controlType, path string) []ValidationIssue {
	if !isPageSelectionControl(controlType) {
		return nil
	}
	value, ok := lookupRawField(values, "selectedIndex")
	if !ok || value == nil || !isInteger(value) {
		return nil // Type validation reports non-integer values.
	}
	index, ok := rawFloat64(value)
	if !ok {
		return nil
	}
	invalid := index < -1
	if strings.EqualFold(strings.TrimSpace(controlType), "TabStrip") {
		if rawTabs, present := lookupRawField(values, "tabs"); present && rawTabs != nil {
			if tabs, array := asSlice(rawTabs); array {
				invalid = len(tabs) == 0 && index != -1 || len(tabs) > 0 && (index < 0 || index >= float64(len(tabs)))
			}
		}
	}
	if !invalid {
		return nil
	}
	field := path + ".selectedIndex"
	return []ValidationIssue{selectedIndexIssue(field)}
}

func rawSelectedIndex(value any) *float64 {
	if value == nil || !isInteger(value) {
		return nil
	}
	index, ok := rawFloat64(value)
	if !ok {
		return nil
	}
	return &index
}

func selectedIndexIssue(path string) ValidationIssue {
	return validationIssue(selectedIndexValidationCode, SeverityError, fmt.Sprintf("%s must be -1 or a valid zero-based index.", path), path, "Use -1 for no selection or an index within the declared collection.", "")
}

func isCustomProgID(progID string) bool {
	trimmed := strings.TrimSpace(progID)
	if trimmed == "" {
		return false
	}
	_, known := LookupControlContractByProgID(trimmed)
	return !known
}

func addGenericCustomControlProperties(properties map[string]PropertyContract, support SupportLevel) {
	properties["caption"] = property(ValueTypeString, false, support, "Generic custom control caption captured from Excel.", false)
	properties["text"] = property(ValueTypeString, false, support, "Generic custom control text captured from Excel.", false)
	properties["value"] = property(ValueTypeAny, false, support, "Generic custom control value captured from Excel.", false)
}

func markUnsupportedControlProperties(issues []ValidationIssue, controlType string) {
	for i := range issues {
		if issues[i].Code != "UFV001" {
			continue
		}
		propertyName := validationFieldName(issues[i].Field)
		if !knownControlPropertyName(propertyName) {
			continue
		}
		issues[i].Code = "UFV005"
		issues[i].Message = fmt.Sprintf("%s is not valid for control type %s.", issues[i].Field, controlType)
		issues[i].Suggestion = "Remove the property or use a control type that supports it."
	}
}

func knownControlPropertyName(name string) bool {
	contract := UserFormContract()
	if _, ok := lookupProperty(contract.CommonControlProperties, name); ok {
		return true
	}
	for _, control := range contract.Controls {
		if _, ok := lookupProperty(control.Properties, name); ok {
			return true
		}
	}
	return false
}

func validationFieldName(field string) string {
	field = strings.TrimSpace(field)
	if dot := strings.LastIndex(field, "."); dot >= 0 {
		return field[dot+1:]
	}
	if bracket := strings.LastIndex(field, "]"); bracket >= 0 && bracket+1 < len(field) && field[bracket+1] == '.' {
		return field[bracket+2:]
	}
	return field
}

func validateObjectProperties(root map[string]any, properties map[string]PropertyContract, path string, allow map[string]bool) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	for _, key := range sortedMapKeys(root) {
		value := root[key]
		field := joinFieldPath(path, key)
		property, ok := lookupProperty(properties, key)
		if !ok {
			if allow != nil && allow[key] {
				continue
			}
			issues = append(issues, validationIssue("UFV001", SeverityError, fmt.Sprintf("%s is not defined by the UserForm spec contract.", field), field, "Remove the field or move custom data under properties.", ""))
			continue
		}
		if !valueMatchesType(value, property.ValueType) {
			issues = append(issues, validationIssue("UFV002", SeverityError, fmt.Sprintf("%s must be %s.", field, property.ValueType), field, "", ""))
			continue
		}
		if len(property.AllowedValues) > 0 && !valueInAllowedValues(value, property.AllowedValues) {
			issues = append(issues, validationIssue("UFV003", SeverityError, fmt.Sprintf("%s must be one of: %s.", field, strings.Join(property.AllowedValues, ", ")), field, "", ""))
		}
		if property.SupportLevel != "" && property.SupportLevel != SupportLevelSupported {
			issues = append(issues, validationIssue("UFV013", SeverityWarning, fmt.Sprintf("%s has %s support.", field, property.SupportLevel), field, supportSuggestion(property.SupportLevel), property.SupportLevel))
		}
	}
	for _, key := range sortedPropertyKeys(properties) {
		property := properties[key]
		if !property.Required {
			continue
		}
		if _, ok := lookupRawField(root, key); !ok {
			issues = append(issues, requiredFieldIssue(joinFieldPath(path, key)))
		}
	}
	return issues
}

func formBuildProperties() map[string]PropertyContract {
	return map[string]PropertyContract{
		"caption":      property(ValueTypeString, false, SupportLevelSupported, "Build caption.", false),
		"width":        property(ValueTypeNumber, false, SupportLevelBestEffort, "Build outer width.", false),
		"height":       property(ValueTypeNumber, false, SupportLevelBestEffort, "Build outer height.", false),
		"clientWidth":  property(ValueTypeNumber, false, SupportLevelSupported, "Build client width for pure-Go generation.", false),
		"clientHeight": property(ValueTypeNumber, false, SupportLevelSupported, "Build client height for pure-Go generation.", false),
	}
}

func validateFormDimensions(form FormSpecForm) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	for _, field := range []struct {
		path  string
		value *float64
	}{
		{path: "form.width", value: form.Width},
		{path: "form.height", value: form.Height},
		{path: "form.observed.width", value: observedFormWidth(form.Observed)},
		{path: "form.observed.height", value: observedFormHeight(form.Observed)},
		{path: "form.observed.insideWidth", value: observedFormInsideWidth(form.Observed)},
		{path: "form.observed.insideHeight", value: observedFormInsideHeight(form.Observed)},
		{path: "form.observed.clientWidth", value: observedFormClientWidth(form.Observed)},
		{path: "form.observed.clientHeight", value: observedFormClientHeight(form.Observed)},
	} {
		if field.value != nil && !validFormDimension(*field.value) {
			issues = append(issues, invalidFormDimensionIssue(field.path))
		}
	}
	if build := form.Build; build != nil {
		for _, field := range []struct {
			path  string
			value *float64
		}{
			{path: "form.build.width", value: build.Width},
			{path: "form.build.height", value: build.Height},
			{path: "form.build.clientWidth", value: build.ClientWidth},
			{path: "form.build.clientHeight", value: build.ClientHeight},
		} {
			if field.value != nil && !validFormDimension(*field.value) {
				issues = append(issues, invalidFormDimensionIssue(field.path))
			}
		}
		if hasBuildClientDimensions(build) && (hasBuildOuterDimensions(build) || form.Width != nil || form.Height != nil) {
			issues = append(issues, validationIssue(
				formBuildDimensionConflictCode,
				SeverityError,
				"legacy form.width/height or form.build.width/height and form.build.clientWidth/clientHeight cannot be specified together.",
				"form.build",
				"Use legacy outer dimensions without client dimensions, or use clientWidth/clientHeight for pure-Go client dimensions.",
				SupportLevelSupported,
			))
		}
	}
	return issues
}

func validateRawFormDimensions(formMap map[string]any) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	legacyOuterDimensions := hasRawFormDimension(formMap, "width") || hasRawFormDimension(formMap, "height")
	for _, field := range []struct {
		root map[string]any
		key  string
		path string
	}{
		{root: formMap, key: "width", path: "form.width"},
		{root: formMap, key: "height", path: "form.height"},
	} {
		issues = append(issues, validateRawFormDimension(field.root, field.key, field.path)...)
	}
	var buildObject map[string]any
	for _, nested := range []struct {
		key  string
		path string
	}{
		{key: "observed", path: "form.observed"},
		{key: "build", path: "form.build"},
	} {
		value, ok := lookupRawField(formMap, nested.key)
		if !ok {
			continue
		}
		object, ok := asObjectMap(value)
		if !ok {
			continue
		}
		if nested.key == "build" {
			buildObject = object
		}
		for _, key := range []string{"width", "height", "insideWidth", "insideHeight", "clientWidth", "clientHeight"} {
			issues = append(issues, validateRawFormDimension(object, key, nested.path+"."+key)...)
		}
	}
	if buildObject != nil && (legacyOuterDimensions || hasRawFormDimension(buildObject, "width") || hasRawFormDimension(buildObject, "height")) && (hasRawFormDimension(buildObject, "clientWidth") || hasRawFormDimension(buildObject, "clientHeight")) {
		issues = append(issues, validationIssue(
			formBuildDimensionConflictCode,
			SeverityError,
			"legacy form.width/height or form.build.width/height and form.build.clientWidth/clientHeight cannot be specified together.",
			"form.build",
			"Use legacy outer dimensions without client dimensions, or use clientWidth/clientHeight for pure-Go client dimensions.",
			SupportLevelSupported,
		))
	}
	return issues
}

func validateRawFormDimension(root map[string]any, key, path string) []ValidationIssue {
	value, ok := lookupRawField(root, key)
	if !ok || value == nil {
		return nil
	}
	if !isNumber(value) {
		return nil
	}
	if numeric, ok := rawFloat64(value); ok && !validFormDimension(numeric) {
		return []ValidationIssue{invalidFormDimensionIssue(path)}
	}
	return nil
}

func hasRawFormDimension(root map[string]any, key string) bool {
	value, ok := lookupRawField(root, key)
	return ok && value != nil
}

func hasBuildOuterDimensions(build *FormSpecBuildForm) bool {
	return build != nil && (build.Width != nil || build.Height != nil)
}

func hasBuildClientDimensions(build *FormSpecBuildForm) bool {
	return build != nil && (build.ClientWidth != nil || build.ClientHeight != nil)
}

func rawFloat64(value any) (float64, bool) {
	switch typed := value.(type) {
	case int:
		return float64(typed), true
	case int8:
		return float64(typed), true
	case int16:
		return float64(typed), true
	case int32:
		return float64(typed), true
	case int64:
		return float64(typed), true
	case uint:
		return float64(typed), true
	case uint8:
		return float64(typed), true
	case uint16:
		return float64(typed), true
	case uint32:
		return float64(typed), true
	case uint64:
		return float64(typed), true
	case float32:
		return float64(typed), true
	case float64:
		return typed, true
	case json.Number:
		numeric, err := typed.Float64()
		return numeric, err == nil
	default:
		return 0, false
	}
}

func validFormDimension(value float64) bool {
	if math.IsNaN(value) || math.IsInf(value, 0) || value < 0 {
		return false
	}
	himetric := math.Round(value * 2540 / 72)
	return !math.IsInf(himetric, 0) && himetric <= float64(math.MaxInt32)
}

func invalidFormDimensionIssue(path string) ValidationIssue {
	return validationIssue(
		formDimensionValidationCode,
		SeverityError,
		fmt.Sprintf("%s must be finite, non-negative, and within the supported HIMETRIC range.", path),
		path,
		"Use a finite dimension that converts to a non-negative signed 32-bit HIMETRIC value.",
		SupportLevelSupported,
	)
}

func formObservedProperties() map[string]PropertyContract {
	return map[string]PropertyContract{
		"caption":      property(ValueTypeString, false, SupportLevelSnapshotOnly, "Observed caption.", false),
		"width":        property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed width.", false),
		"height":       property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed height.", false),
		"insideWidth":  property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed inside width.", false),
		"insideHeight": property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed inside height.", false),
		"clientWidth":  property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed client width.", false),
		"clientHeight": property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed client height.", false),
	}
}

func parentCycleIssues(parentByID, parentFieldByID map[string]string) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	reported := map[string]bool{}
	for _, id := range sortedMapKeys(parentByID) {
		if reported[id] {
			continue
		}
		path := make([]string, 0)
		seen := make(map[string]int)
		for current := id; current != ""; current = parentByID[current] {
			if cycleStart, exists := seen[current]; exists {
				cycle := path[cycleStart:]
				for _, cycleID := range cycle {
					reported[cycleID] = true
				}
				field := parentFieldByID[cycle[0]]
				if field == "" {
					field = "controls"
				}
				issues = append(issues, validationIssue("UFV010", SeverityError, fmt.Sprintf("controls parentId chain for %q contains a cycle.", cycle[0]), field, "Break the parentId cycle so controls form a tree.", ""))
				break
			}
			if reported[current] {
				break
			}
			seen[current] = len(path)
			path = append(path, current)
		}
	}
	return issues
}

func sortedMapKeys[T any](values map[string]T) []string {
	keys := make([]string, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

func sortedPropertyKeys(values map[string]PropertyContract) []string {
	return sortedMapKeys(values)
}

func valueMatchesType(value any, valueType ValueType) bool {
	if value == nil {
		return true
	}
	switch valueType {
	case ValueTypeAny:
		return true
	case ValueTypeString:
		_, ok := value.(string)
		return ok
	case ValueTypeNumber:
		return isNumber(value)
	case ValueTypeInteger:
		return isInteger(value)
	case ValueTypeBoolean:
		_, ok := value.(bool)
		return ok
	case ValueTypeStringArray:
		items, ok := asSlice(value)
		if !ok {
			return false
		}
		for _, item := range items {
			if _, ok := item.(string); !ok {
				return false
			}
		}
		return true
	case ValueTypeObject:
		_, ok := asObjectMap(value)
		return ok
	case ValueTypeObjectArray:
		items, ok := asSlice(value)
		if !ok {
			return false
		}
		for _, item := range items {
			if _, ok := asObjectMap(item); !ok {
				return false
			}
		}
		return true
	default:
		return false
	}
}

func isNumber(value any) bool {
	switch typed := value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		return true
	case float32:
		return !math.IsNaN(float64(typed)) && !math.IsInf(float64(typed), 0)
	case float64:
		return !math.IsNaN(typed) && !math.IsInf(typed, 0)
	case json.Number:
		numeric, err := typed.Float64()
		return err == nil && !math.IsNaN(numeric) && !math.IsInf(numeric, 0)
	default:
		return false
	}
}

func isInteger(value any) bool {
	switch typed := value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		return true
	case float32:
		return !math.IsNaN(float64(typed)) && !math.IsInf(float64(typed), 0) && math.Trunc(float64(typed)) == float64(typed)
	case float64:
		return !math.IsNaN(typed) && !math.IsInf(typed, 0) && math.Trunc(typed) == typed
	case json.Number:
		_, err := typed.Int64()
		return err == nil
	default:
		return false
	}
}

func valueInAllowedValues(value any, allowed []string) bool {
	var text string
	switch typed := value.(type) {
	case string:
		text = typed
	case json.Number:
		text = typed.String()
	default:
		text = fmt.Sprint(typed)
	}
	for _, item := range allowed {
		if text == item {
			return true
		}
	}
	return false
}

func lookupRawField(root map[string]any, key string) (any, bool) {
	if value, ok := root[key]; ok {
		return value, true
	}
	normalized := contractKey(key)
	for rawKey, value := range root {
		if contractKey(rawKey) == normalized {
			return value, true
		}
	}
	return nil, false
}

func normalizeDecodedValue(value any) any {
	switch typed := value.(type) {
	case map[string]any:
		out := make(map[string]any, len(typed))
		for key, item := range typed {
			out[key] = normalizeDecodedValue(item)
		}
		return out
	case map[any]any:
		out := make(map[string]any, len(typed))
		for key, item := range typed {
			out[fmt.Sprint(key)] = normalizeDecodedValue(item)
		}
		return out
	case []any:
		out := make([]any, 0, len(typed))
		for _, item := range typed {
			out = append(out, normalizeDecodedValue(item))
		}
		return out
	default:
		return value
	}
}

func validationIssue(code string, severity Severity, message, field, suggestion string, support SupportLevel) ValidationIssue {
	return ValidationIssue{
		Code:       code,
		Severity:   severity,
		Message:    message,
		Field:      field,
		Suggestion: suggestion,
		Support:    support,
	}
}

func invalidFixedValueIssue(field, want string) ValidationIssue {
	return validationIssue("UFV003", SeverityError, fmt.Sprintf("%s must be %s.", field, want), field, "", "")
}

func requiredFieldIssue(field string) ValidationIssue {
	return validationIssue("UFV004", SeverityError, fmt.Sprintf("%s is required.", field), field, "", "")
}

func supportSuggestion(level SupportLevel) string {
	switch level {
	case SupportLevelBestEffort:
		return "Treat this field as best-effort and verify the rebuilt form in Excel."
	case SupportLevelObservedOnly:
		return "Treat this field as observed state; xlflow may apply it best-effort but does not guarantee round-trip fidelity."
	case SupportLevelSnapshotOnly:
		return "This field is snapshot-oriented and should not be treated as authoritative build intent."
	case SupportLevelCustomUnchecked:
		return "Custom controls are accepted, but xlflow cannot validate type-specific behavior."
	default:
		return ""
	}
}

func joinFieldPath(base, key string) string {
	if base == "" {
		return key
	}
	return base + "." + key
}

func hasValidationErrors(issues []ValidationIssue) bool {
	return firstValidationError(issues).Code != ""
}

func firstValidationError(issues []ValidationIssue) ValidationIssue {
	for _, issue := range issues {
		if issue.Severity == SeverityError {
			return issue
		}
	}
	return ValidationIssue{}
}

func validationWarnings(issues []ValidationIssue) []ValidationIssue {
	warnings := make([]ValidationIssue, 0)
	seen := map[string]bool{}
	for _, issue := range issues {
		if issue.Severity != SeverityWarning {
			continue
		}
		key := issue.Code + "\x00" + issue.Field + "\x00" + issue.Message
		if seen[key] {
			continue
		}
		seen[key] = true
		warnings = append(warnings, issue)
	}
	return warnings
}

func newSpecValidationIssuesError(input SpecInput, issues []ValidationIssue) error {
	err := newSpecValidationErrorFromIssue(firstValidationError(issues), issues)
	if err.Path == "" {
		err.Path = input.DisplayPath
	}
	if err.Format == "" {
		err.Format = input.Format
	}
	return err
}

func newSpecValidationErrorFromIssue(issue ValidationIssue, issues []ValidationIssue) *SpecError {
	code := "spec_validation_failed"
	if issue.Code == "UFV003" || (issue.Code == "UFV004" && (issue.Field == "schemaVersion" || issue.Field == "kind" || issue.Field == "basis" || issue.Field == "form" || issue.Field == "form.name")) {
		code = "spec_schema_invalid"
	}
	message := issue.Message
	if message == "" {
		message = "UserForm spec validation failed"
	}
	return &SpecError{
		Code:       code,
		Message:    message,
		Field:      issue.Field,
		Suggestion: issue.Suggestion,
		Issues:     append([]ValidationIssue(nil), issues...),
	}
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
	if !hasBuildClientDimensions(form.Build) && form.Build.Width == nil {
		form.Build.Width = firstFloatPtr(form.Width, observedFormWidth(form.Observed))
	}
	if !hasBuildClientDimensions(form.Build) && form.Build.Height == nil {
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
	if control.Observed.Tag == nil && control.Tag != nil {
		control.Observed.Tag = control.Tag
	}
	if control.Observed.ControlTipText == nil && control.ControlTipText != nil {
		control.Observed.ControlTipText = control.ControlTipText
	}
	if control.Observed.Accelerator == nil && control.Accelerator != nil {
		control.Observed.Accelerator = control.Accelerator
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
	if control.Observed.Tabs == nil && control.Tabs != nil {
		control.Observed.Tabs = slices.Clone(control.Tabs)
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
	properties, _ := asObjectMap(root["properties"])
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
	if tag, ok := snapshotStringField(root, properties, "tag"); ok {
		control.Tag = &tag
	}
	if controlTipText, ok := snapshotStringField(root, properties, "control_tip_text", "controlTipText"); ok {
		control.ControlTipText = &controlTipText
	}
	if accelerator, ok := snapshotStringField(root, properties, "accelerator"); ok {
		control.Accelerator = &accelerator
	}
	if text, ok := stringField(root, "text"); ok {
		control.Text = &text
	}
	if value, ok := root["value"]; ok && value != nil {
		_, knownType := LookupControlContract(controlType)
		_, supportsValue := LookupControlProperty(controlType, "value")
		if supportsValue || (!knownType && strings.TrimSpace(control.ProgID) != "") {
			control.Value = value
		} else {
			control.Unsupported = append(control.Unsupported, "value")
			warnings = append(warnings, FormSpecWarning{
				Code:    "unsupported_properties",
				Message: "Unsupported Designer properties were omitted from the FormSpec snapshot: value.",
				Control: control.Name,
			})
		}
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
	if rawTabs, found := lookupRawField(root, "tabs"); !found {
		if rawTabs, found = lookupRawField(properties, "tabs"); found {
			tabs, tabsErr := formSpecTabs(rawTabs)
			if tabsErr != nil {
				return FormSpecControl{}, nil, nil, tabsErr
			}
			control.Tabs = tabs
		}
	} else {
		tabs, tabsErr := formSpecTabs(rawTabs)
		if tabsErr != nil {
			return FormSpecControl{}, nil, nil, tabsErr
		}
		control.Tabs = tabs
	}
	if unsupported, ok := stringSliceField(root, "unsupported"); ok {
		control.Unsupported = append(control.Unsupported, unsupported...)
		slices.Sort(control.Unsupported)
		control.Unsupported = slices.Compact(control.Unsupported)
	}
	if len(properties) > 0 {
		control.Properties = properties
	}
	if strings.EqualFold(strings.TrimSpace(control.Type), "Page") &&
		(control.Left != nil || control.Top != nil || control.Width != nil || control.Height != nil) {
		control.Observed = &FormSpecObservedControl{
			Left:   control.Left,
			Top:    control.Top,
			Width:  control.Width,
			Height: control.Height,
		}
		control.Left = nil
		control.Top = nil
		control.Width = nil
		control.Height = nil
	}
	control = normalizeObservedControl(control)
	children, childWarnings, err := formSpecControls(root["controls"], control.ID, unnamedCounter, idCounter)
	if err != nil {
		return FormSpecControl{}, nil, nil, err
	}
	warnings = append(warnings, childWarnings...)
	return control, children, warnings, nil
}

func snapshotStringField(root, properties map[string]any, keys ...string) (string, bool) {
	for _, source := range []map[string]any{root, properties} {
		for _, key := range keys {
			if value, ok := stringField(source, key); ok {
				return value, true
			}
		}
	}
	return "", false
}

func formSpecTabs(value any) ([]FormSpecTab, error) {
	if value == nil {
		return nil, nil
	}
	items, ok := asSlice(value)
	if !ok {
		return nil, fmt.Errorf("inspect designer snapshot tabs value is invalid")
	}
	tabs := make([]FormSpecTab, 0, len(items))
	for index, item := range items {
		object, ok := asObjectMap(item)
		if !ok {
			return nil, fmt.Errorf("inspect designer snapshot tab entry %d is invalid", index)
		}
		tab := FormSpecTab{}
		if tab.Name, ok = stringField(object, "name"); !ok {
			return nil, fmt.Errorf("inspect designer snapshot tab entry %d is missing a name", index)
		}
		if caption, ok := stringField(object, "caption"); ok {
			tab.Caption = &caption
		}
		if controlTipText, ok := snapshotStringField(object, nil, "control_tip_text", "controlTipText"); ok {
			tab.ControlTipText = &controlTipText
		}
		if tag, ok := stringField(object, "tag"); ok {
			tab.Tag = &tag
		}
		if accelerator, ok := stringField(object, "accelerator"); ok {
			tab.Accelerator = &accelerator
		}
		if enabled, ok := optionalBoolField(object, "enabled"); ok {
			tab.Enabled = &enabled
		}
		if visible, ok := optionalBoolField(object, "visible"); ok {
			tab.Visible = &visible
		}
		tabs = append(tabs, tab)
	}
	return tabs, nil
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
	return observed != nil && (observed.Caption != nil || observed.Width != nil || observed.Height != nil || observed.InsideWidth != nil || observed.InsideHeight != nil || observed.ClientWidth != nil || observed.ClientHeight != nil)
}

func hasBuildFormValues(build *FormSpecBuildForm) bool {
	return build != nil && (build.Caption != nil || build.Width != nil || build.Height != nil || build.ClientWidth != nil || build.ClientHeight != nil)
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

func observedFormInsideWidth(observed *FormSpecObservedForm) *float64 {
	if observed == nil {
		return nil
	}
	return observed.InsideWidth
}

func observedFormInsideHeight(observed *FormSpecObservedForm) *float64 {
	if observed == nil {
		return nil
	}
	return observed.InsideHeight
}

func observedFormClientWidth(observed *FormSpecObservedForm) *float64 {
	if observed == nil {
		return nil
	}
	return observed.ClientWidth
}

func observedFormClientHeight(observed *FormSpecObservedForm) *float64 {
	if observed == nil {
		return nil
	}
	return observed.ClientHeight
}

func hasObservedControlValues(observed *FormSpecObservedControl) bool {
	return observed != nil &&
		(observed.Caption != nil ||
			observed.Tag != nil ||
			observed.ControlTipText != nil ||
			observed.Accelerator != nil ||
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
			observed.Tabs != nil ||
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
