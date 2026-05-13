package excel

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"gopkg.in/yaml.v3"
)

type FormSnapshotOutput struct {
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
	Name    string   `json:"name" yaml:"name"`
	Caption *string  `json:"caption,omitempty" yaml:"caption,omitempty"`
	Width   *float64 `json:"width,omitempty" yaml:"width,omitempty"`
	Height  *float64 `json:"height,omitempty" yaml:"height,omitempty"`
}

type FormSpecControl struct {
	Type          string            `json:"type" yaml:"type"`
	Name          string            `json:"name" yaml:"name"`
	ProgID        string            `json:"progId,omitempty" yaml:"progId,omitempty"`
	Caption       *string           `json:"caption,omitempty" yaml:"caption,omitempty"`
	Text          *string           `json:"text,omitempty" yaml:"text,omitempty"`
	Value         any               `json:"value,omitempty" yaml:"value,omitempty"`
	Left          *float64          `json:"left,omitempty" yaml:"left,omitempty"`
	Top           *float64          `json:"top,omitempty" yaml:"top,omitempty"`
	Width         *float64          `json:"width,omitempty" yaml:"width,omitempty"`
	Height        *float64          `json:"height,omitempty" yaml:"height,omitempty"`
	TabIndex      *int              `json:"tabIndex,omitempty" yaml:"tabIndex,omitempty"`
	SelectedIndex *int              `json:"selectedIndex,omitempty" yaml:"selectedIndex,omitempty"`
	Enabled       *bool             `json:"enabled,omitempty" yaml:"enabled,omitempty"`
	Visible       *bool             `json:"visible,omitempty" yaml:"visible,omitempty"`
	List          []string          `json:"list,omitempty" yaml:"list,omitempty"`
	Unsupported   []string          `json:"unsupported,omitempty" yaml:"unsupported,omitempty"`
	Controls      []FormSpecControl `json:"controls,omitempty" yaml:"controls,omitempty"`
	Properties    map[string]any    `json:"properties,omitempty" yaml:"properties,omitempty"`
}

type FormSpecWarning struct {
	Code    string `json:"code,omitempty" yaml:"code,omitempty"`
	Message string `json:"message,omitempty" yaml:"message,omitempty"`
	Control string `json:"control,omitempty" yaml:"control,omitempty"`
}

func ResolveFormSnapshotOutput(root, outPath string) (FormSnapshotOutput, error) {
	trimmed := strings.TrimSpace(outPath)
	if trimmed == "" {
		return FormSnapshotOutput{}, fmt.Errorf("--out is required")
	}
	resolved := trimmed
	if !filepath.IsAbs(resolved) {
		resolved = filepath.Join(root, resolved)
	}
	resolved = filepath.Clean(resolved)
	format, err := formSnapshotFormatFromPath(resolved)
	if err != nil {
		return FormSnapshotOutput{}, err
	}
	if info, statErr := os.Stat(resolved); statErr == nil && info.IsDir() {
		return FormSnapshotOutput{}, fmt.Errorf("output path %q is a directory", trimmed)
	} else if statErr != nil && !os.IsNotExist(statErr) {
		return FormSnapshotOutput{}, statErr
	}
	return FormSnapshotOutput{
		Path:        resolved,
		DisplayPath: filepath.ToSlash(relPath(root, resolved)),
		Format:      format,
	}, nil
}

func WriteFormSnapshot(output FormSnapshotOutput, spec FormSpec) error {
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
	counter := 0
	controls, generatedWarnings, err := formSpecControls(root["controls"], &counter)
	if err != nil {
		return FormSpec{}, err
	}
	warnings, err := formSpecWarnings(root["warnings"])
	if err != nil {
		return FormSpec{}, err
	}
	warnings = append(warnings, generatedWarnings...)
	form := FormSpecForm{Name: name}
	if caption, ok := stringField(root, "caption"); ok {
		form.Caption = &caption
	}
	if width, ok := optionalFloatField(root, "width"); ok {
		form.Width = &width
	}
	if height, ok := optionalFloatField(root, "height"); ok {
		form.Height = &height
	}
	if warnings == nil {
		warnings = []FormSpecWarning{}
	}
	return FormSpec{
		SchemaVersion:    1,
		Kind:             "xlflow.userform",
		Basis:            basis,
		CoordinateSystem: coordinateSystem,
		Form:             form,
		Controls:         controls,
		Warnings:         warnings,
	}, nil
}

func formSnapshotFormatFromPath(path string) (string, error) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".json":
		return "json", nil
	case ".yaml", ".yml":
		return "yaml", nil
	default:
		return "", fmt.Errorf("--out must end with .json, .yaml, or .yml")
	}
}

func formSpecControls(value any, unnamedCounter *int) ([]FormSpecControl, []FormSpecWarning, error) {
	items, ok := asSlice(value)
	if !ok || len(items) == 0 {
		return []FormSpecControl{}, []FormSpecWarning{}, nil
	}
	controls := make([]FormSpecControl, 0, len(items))
	warnings := make([]FormSpecWarning, 0)
	for _, item := range items {
		controlMap, ok := asObjectMap(item)
		if !ok {
			return nil, nil, fmt.Errorf("inspect designer snapshot control entry is invalid")
		}
		control, controlWarnings, err := formSpecControl(controlMap, unnamedCounter)
		if err != nil {
			return nil, nil, err
		}
		controls = append(controls, control)
		warnings = append(warnings, controlWarnings...)
	}
	return controls, warnings, nil
}

func formSpecControl(root map[string]any, unnamedCounter *int) (FormSpecControl, []FormSpecWarning, error) {
	name, ok := stringField(root, "name")
	warnings := make([]FormSpecWarning, 0)
	if !ok || strings.TrimSpace(name) == "" {
		if unnamedCounter == nil {
			return FormSpecControl{}, nil, fmt.Errorf("inspect designer snapshot control is missing a name")
		}
		*unnamedCounter++
		name = fmt.Sprintf("<unnamed_%d>", *unnamedCounter)
		warnings = append(warnings, FormSpecWarning{
			Code:    "unnamed_control_placeholder",
			Message: "A control without a stable name was persisted with a generated placeholder name.",
			Control: name,
		})
	}
	controlType, ok := stringField(root, "type")
	if !ok || strings.TrimSpace(controlType) == "" {
		return FormSpecControl{}, nil, fmt.Errorf("inspect designer snapshot control %q is missing a type", name)
	}
	children, childWarnings, err := formSpecControls(root["controls"], unnamedCounter)
	if err != nil {
		return FormSpecControl{}, nil, err
	}
	warnings = append(warnings, childWarnings...)
	control := FormSpecControl{
		Name:     name,
		Type:     controlType,
		Controls: children,
	}
	if progID, ok := stringField(root, "prog_id"); ok {
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
	return control, warnings, nil
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
