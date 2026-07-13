package forms

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
	for i, control := range spec.Controls {
		if strings.TrimSpace(control.ParentID) == "" {
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
		if canContainChildren, knownParentControl := FormSpecControlCanContainChildren(parent); knownParentControl && !canContainChildren {
			issues = append(issues, validationIssue("UFV011", SeverityError, fmt.Sprintf("%s %q references non-container control %q.", field, control.ParentID, parent.Name), field, "Use a Frame or custom container control as the parent.", ""))
		}
		parentByID[control.ID] = control.ParentID
	}
	issues = append(issues, parentCycleIssues(parentByID)...)
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
	if _, err := ControlProgID(control); err != nil {
		issues = append(issues, validationIssue("UFV006", SeverityError, fmt.Sprintf("%s: %v.", path, err), path+".type", "Use a supported built-in type or provide a custom progId.", ""))
	}
	if strings.TrimSpace(control.Type) != "" && strings.TrimSpace(control.ProgID) != "" {
		if progControl, ok := LookupControlContractByProgID(control.ProgID); ok && !strings.EqualFold(strings.TrimSpace(control.Type), progControl.Type) {
			issues = append(issues, validationIssue("UFV012", SeverityError, fmt.Sprintf("%s.progId %q is for %s, not %s.", path, control.ProgID, progControl.Type, control.Type), path+".progId", "Use the ProgID that matches type or change type to match the ProgID.", ""))
		}
	}
	return issues
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
	Path     string
	ID       string
	ParentID string
	Name     string
	Type     string
	ProgID   string
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
	refs := []rawControlRef{ref}
	if observed, ok := asObjectMap(controlMap["observed"]); ok {
		issues = append(issues, validateRawObservedControlProperties(observed, ref.Type, ref.ProgID, path+".observed")...)
	}
	if children, ok := controlMap["controls"]; ok {
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
	builtInControl, builtInType := LookupControlContract(controlType)
	if builtInType {
		for key, property := range builtInControl.Properties {
			allowed[key] = property
		}
	}
	issues := validateObjectProperties(controlMap, allowed, path, nil)
	if builtInType {
		markUnsupportedControlProperties(issues, controlType)
	}
	if strings.TrimSpace(controlType) != "" && strings.TrimSpace(progID) != "" {
		if progControl, ok := LookupControlContractByProgID(progID); ok && !strings.EqualFold(strings.TrimSpace(controlType), progControl.Type) {
			issues = append(issues, validationIssue("UFV012", SeverityError, fmt.Sprintf("%s.progId %q is for %s, not %s.", path, progID, progControl.Type, controlType), path+".progId", "Use the ProgID that matches type or change type to match the ProgID.", ""))
		}
	}
	if strings.TrimSpace(controlType) != "" && !builtInType {
		if strings.TrimSpace(progID) == "" {
			issues = append(issues, validationIssue("UFV006", SeverityError, fmt.Sprintf("%s.type %q is not a supported built-in control type.", path, controlType), path+".type", "Use a supported built-in type or provide a custom progId.", ""))
		} else if _, knownProgID := LookupControlContractByProgID(progID); !knownProgID {
			issues = append(issues, validationIssue("UFV014", SeverityWarning, fmt.Sprintf("%s uses custom control type %q with unchecked ProgID %q.", path, controlType, progID), path+".progId", "Only common structural fields and the properties bag are validated for custom controls.", SupportLevelCustomUnchecked))
		}
	}
	return issues
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
	} else if strings.TrimSpace(progID) == "" {
		allowed["caption"] = property(ValueTypeString, false, SupportLevelSupported, "Observed caption.", false)
		allowed["text"] = property(ValueTypeString, false, SupportLevelSupported, "Observed text.", false)
		allowed["value"] = property(ValueTypeAny, false, SupportLevelSupported, "Observed value.", false)
	}
	issues := validateObjectProperties(controlMap, allowed, path, nil)
	if _, ok := LookupControlContract(controlType); ok {
		markUnsupportedControlProperties(issues, controlType)
	}
	return issues
}

func validateRawControlStructure(controls []rawControlRef) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	ids := make(map[string]rawControlRef, len(controls))
	parentByID := make(map[string]string, len(controls))
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
		if canContainChildren, known := rawControlCanContainChildren(parent); known && !canContainChildren {
			issues = append(issues, validationIssue("UFV011", SeverityError, fmt.Sprintf("%s %q references non-container control %q.", field, control.ParentID, parent.Name), field, "Use a Frame or custom container control as the parent.", ""))
		}
		if control.ID != "" {
			parentByID[control.ID] = control.ParentID
		}
	}
	return append(issues, parentCycleIssues(parentByID)...)
}

func rawControlCanContainChildren(control rawControlRef) (bool, bool) {
	if control.ProgID != "" {
		if contract, ok := LookupControlContractByProgID(control.ProgID); ok {
			return contract.CanContainChildren, true
		}
		return false, false
	}
	if contract, ok := LookupControlContract(control.Type); ok {
		return contract.CanContainChildren, true
	}
	return false, false
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
	for key, value := range root {
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
	for key, property := range properties {
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
		"caption": property(ValueTypeString, false, SupportLevelSupported, "Build caption.", false),
		"width":   property(ValueTypeNumber, false, SupportLevelBestEffort, "Build width.", false),
		"height":  property(ValueTypeNumber, false, SupportLevelBestEffort, "Build height.", false),
	}
}

func formObservedProperties() map[string]PropertyContract {
	return map[string]PropertyContract{
		"caption":      property(ValueTypeString, false, SupportLevelSnapshotOnly, "Observed caption.", false),
		"width":        property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed width.", false),
		"height":       property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed height.", false),
		"insideWidth":  property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed inside width.", false),
		"insideHeight": property(ValueTypeNumber, false, SupportLevelSnapshotOnly, "Observed inside height.", false),
	}
}

func parentCycleIssues(parentByID map[string]string) []ValidationIssue {
	issues := make([]ValidationIssue, 0)
	reported := map[string]bool{}
	for id := range parentByID {
		seen := map[string]bool{}
		for current := id; current != ""; current = parentByID[current] {
			if seen[current] {
				if !reported[id] {
					issues = append(issues, validationIssue("UFV010", SeverityError, fmt.Sprintf("controls parentId chain for %q contains a cycle.", id), "controls", "Break the parentId cycle so controls form a tree.", ""))
					reported[id] = true
				}
				break
			}
			seen[current] = true
		}
	}
	return issues
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
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64:
		return true
	case json.Number:
		_, err := typed.Float64()
		return err == nil
	default:
		return false
	}
}

func isInteger(value any) bool {
	switch typed := value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		return true
	case float32:
		return math.Trunc(float64(typed)) == float64(typed)
	case float64:
		return math.Trunc(typed) == typed
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
