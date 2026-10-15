package edit

import (
	"bytes"
	"encoding/json"
	"maps"
	"reflect"
	"slices"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

// MarshalJSON retains explicit null while omitting a missing value.
func (op Operation) MarshalJSON() ([]byte, error) {
	type wire Operation
	var value json.RawMessage
	if op.ValuePresent || op.Value != nil {
		var err error
		value, err = json.Marshal(op.Value)
		if err != nil {
			return nil, err
		}
	}
	return json.Marshal(struct {
		wire
		Value json.RawMessage `json:"value,omitempty"`
	}{wire: wire(op), Value: value})
}

// UnmarshalJSON records value presence independently of its decoded value.
func (op *Operation) UnmarshalJSON(data []byte) error {
	type wire Operation
	var decoded struct {
		wire
		Value json.RawMessage `json:"value"`
	}
	if err := json.Unmarshal(data, &decoded); err != nil {
		return err
	}
	next := Operation(decoded.wire)
	next.ValuePresent = decoded.Value != nil
	if next.ValuePresent {
		decoder := json.NewDecoder(bytes.NewReader(decoded.Value))
		decoder.UseNumber()
		if err := decoder.Decode(&next.Value); err != nil {
			return err
		}
	}
	*op = next
	return nil
}

// EditableFormProperties returns exact source paths supported by scalar edits.
func EditableFormProperties() map[string]spec.PropertyContract {
	result := map[string]spec.PropertyContract{}
	for name, contract := range spec.UserFormContract().FormProperties {
		if contract.IncludeInAuthoring && scalarContract(contract) && !editorPropertyName(name) && hasYAMLField(reflect.TypeFor[spec.FormSpecForm](), name) {
			result[name] = contract
		}
	}
	return result
}

// EditableControlProperties excludes topology, snapshots, collections and fields
// which the typed authoring DTO cannot retain.
func EditableControlProperties(typeName string) map[string]spec.PropertyContract {
	result := map[string]spec.PropertyContract{}
	for name, contract := range spec.ControlProperties(typeName) {
		if slices.Contains([]string{"id", "type", "parentId", "zIndex"}, name) {
			continue
		}
		if contract.IncludeInAuthoring && scalarContract(contract) && !editorPropertyName(name) && hasYAMLField(reflect.TypeFor[spec.FormSpecControl](), name) {
			result[name] = contract
		}
	}
	return result
}

// EditableTabProperties returns scalar authoring fields represented by FormSpecTab.
func EditableTabProperties() map[string]spec.PropertyContract {
	result := map[string]spec.PropertyContract{}
	for name, contract := range spec.UserFormContract().TabProperties {
		if contract.IncludeInAuthoring && scalarContract(contract) && !editorPropertyName(name) && hasYAMLField(reflect.TypeFor[spec.FormSpecTab](), name) {
			result[name] = contract
		}
	}
	return result
}

func checkPropertyValue(op Operation, contract spec.PropertyContract) error {
	if op.Value == nil && !op.ValuePresent {
		return failure(-1, op, "UFE003", "A property value is required.", "Supply value; use explicit null only for an optional field.")
	}
	if op.Value == nil && contract.Required {
		return failure(-1, op, "UFE003", "A required property cannot be null.", "Supply a non-null scalar value.")
	}
	if number, ok := op.Value.(json.Number); ok {
		// Int64 payloads retain their exact integer representation in both
		// writers. Other numbers pass through float64 when writing YAML.
		if _, err := number.Int64(); err != nil && !propertyNumberRoundTrips(number) {
			return failure(-1, op, "UFE003", "The number cannot be written without precision loss.", "Use the source editor for this value.")
		}
	}
	if typedPropertyContract(op.Field, contract) {
		return checkTypedValue(op, contract)
	}
	_, err := scalar(op.Value)
	return err
}

type PropertyDescriptor struct {
	Field            string               `json:"field"`
	ValueType        spec.ValueType       `json:"valueType"`
	Required         bool                 `json:"required"`
	Nullable         bool                 `json:"nullable"`
	AllowedValues    []string             `json:"allowedValues,omitempty"`
	Editor           string               `json:"editor,omitempty"`
	Children         []PropertyDescriptor `json:"children,omitempty"`
	Suggestions      []PropertySuggestion `json:"suggestions,omitempty"`
	Minimum          *float64             `json:"minimum,omitempty"`
	Maximum          *float64             `json:"maximum,omitempty"`
	ExclusiveMinimum bool                 `json:"exclusiveMinimum,omitzero"`
}

// Suggestions are named hints, never a restriction on accepted values.
type PropertySuggestion struct {
	Value any    `json:"value"`
	Label string `json:"label"`
}

type PropertyState struct {
	Present bool `json:"present"`
	Legacy  bool `json:"legacy,omitzero"`
	Value   any  `json:"value"`
}

type PropertyTarget struct {
	Descriptors []PropertyDescriptor     `json:"descriptors"`
	Values      map[string]PropertyState `json:"values"`
}

type PropertyGridData struct {
	Form     PropertyTarget              `json:"form"`
	Controls map[string]PropertyTarget   `json:"controls"`
	Tabs     map[string][]PropertyTarget `json:"tabs,omitempty"`
}

func propertyTarget(properties map[string]spec.PropertyContract, lookup func(string) (any, bool, error), editorVersion ...bool) (PropertyTarget, error) {
	target := PropertyTarget{Descriptors: []PropertyDescriptor{}, Values: map[string]PropertyState{}}
	for _, name := range slices.Sorted(maps.Keys(properties)) {
		contract := properties[name]
		value, present, err := lookup(name)
		if err != nil {
			return PropertyTarget{}, err
		}
		legacy := false
		if !present && legacyPropertyName(name) != "" {
			for _, alias := range []string{legacyPropertyName(name), name} {
				if alias == "" {
					continue
				}
				raw, exists, lookupErr := lookup("properties." + alias)
				if lookupErr != nil {
					return PropertyTarget{}, lookupErr
				}
				if exists {
					value, err = legacyEditorValue(name, raw)
					if err != nil {
						return PropertyTarget{}, err
					}
					present, legacy = true, true
					break
				}
			}
		}
		if present {
			if err := checkMetadataValue(value, contract); err != nil {
				return PropertyTarget{}, failure(-1, Operation{Field: name}, "UFE005", "Authored property cannot be represented by this editor.", "Use the source editor for this property.")
			}
			if !propertyNumberRoundTrips(value) {
				return PropertyTarget{}, propertyPrecisionError(name)
			}
		}
		descriptor := propertyDescriptor(name, contract)
		if len(editorVersion) == 0 || !editorVersion[0] {
			descriptor.Editor = ""
			descriptor.Children = nil
		}
		target.Descriptors = append(target.Descriptors, descriptor)
		target.Values[name] = PropertyState{Present: present, Legacy: legacy, Value: value}
	}
	return target, nil
}

func propertyPrecisionError(field string) error {
	return failure(-1, Operation{Field: field}, "UFE005", "Authored number cannot be represented without precision loss in the Property Grid.", "Use the source editor for this property.")
}

// Compare decimal values before and after the JavaScript number wire boundary.
// Ordinary decimals such as 0.1 remain usable; extra authored precision does not.
func propertyNumberRoundTrips(value any) bool {
	switch value.(type) {
	case json.Number, int, int64, uint64, float64:
		encoded, err := json.Marshal(value)
		if err != nil {
			return false
		}
		number, err := strconv.ParseFloat(string(encoded), 64)
		if err != nil {
			return false
		}
		original := rational(json.Number(encoded))
		roundtrip := rational(json.Number(strconv.FormatFloat(number, 'g', -1, 64)))
		return original != nil && roundtrip != nil && original.Cmp(roundtrip) == 0
	default:
		return true
	}
}

// PropertyGrid reports authored presence and scalar values from original syntax,
// never values filled in by FormSpec normalization or observed snapshots.
func PropertyGrid(input spec.SpecInput, source []byte, editorVersion ...int) (PropertyGridData, error) {
	model, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return PropertyGridData{}, validationError(-1, Operation{}, err)
	}
	extended := len(editorVersion) > 0 && editorVersion[0] == 1
	formProperties := EditableFormProperties()
	if extended {
		formProperties = EditableFormEditorProperties()
	}
	result := PropertyGridData{Controls: map[string]PropertyTarget{}}
	if input.Format == "json" {
		root, err := parseJSONSource(source)
		if err != nil {
			return PropertyGridData{}, validationError(-1, Operation{}, err)
		}
		lookup := func(object *jsonSourceNode) func(string) (any, bool, error) {
			return func(path string) (any, bool, error) {
				node := object
				for part := range strings.SplitSeq(path, ".") {
					node = jsonSourceMemberValue(node, part)
				}
				if node == nil {
					return nil, false, nil
				}
				var value any
				decoder := json.NewDecoder(bytes.NewReader(source[node.start:node.end]))
				decoder.UseNumber()
				err := decoder.Decode(&value)
				return value, true, err
			}
		}
		result.Form, err = propertyTarget(formProperties, lookup(jsonSourceMemberValue(root, "form")), extended)
		if err != nil {
			return PropertyGridData{}, err
		}
		nodes := map[string][]*jsonSourceNode{}
		indexJSONControls(jsonSourceMemberValue(root, "controls"), nodes)
		for _, control := range model.Controls {
			if len(nodes[control.ID]) != 1 {
				continue
			} // Inferred IDs are not editable source targets.
			target, err := propertyTarget(editablePropertiesForControl(control, extended), lookup(nodes[control.ID][0]), extended)
			if err != nil {
				return PropertyGridData{}, err
			}
			result.Controls[control.ID] = target
			if strings.EqualFold(control.Type, "TabStrip") {
				tabs := jsonSourceMemberValue(nodes[control.ID][0], "tabs")
				targets := make([]PropertyTarget, 0)
				if tabs != nil && tabs.kind == jsonSourceArray {
					targets = make([]PropertyTarget, 0, len(tabs.elements))
					for _, tab := range tabs.elements {
						target, err := propertyTarget(EditableTabProperties(), lookup(tab))
						if err != nil {
							return PropertyGridData{}, err
						}
						targets = append(targets, target)
					}
				}
				if result.Tabs == nil {
					result.Tabs = map[string][]PropertyTarget{}
				}
				result.Tabs[control.ID] = targets
			}
		}
		return result, nil
	}
	if input.Format != "yaml" {
		return PropertyGridData{}, failure(-1, Operation{}, "UFE001", "Unsupported FormSpec source format.", "Use YAML/YML or JSON.")
	}
	engine, err := newEngine(source)
	if err != nil {
		return PropertyGridData{}, validationError(-1, Operation{}, err)
	}
	// Metadata must not turn alias/merge-dependent fields into absent states.
	// This matches the source writer's conservative refusal of unsafe syntax.
	if unsafe(engine.root) {
		return PropertyGridData{}, unsupported()
	}
	lookup := func(object *yaml.Node) func(string) (any, bool, error) {
		return func(path string) (any, bool, error) {
			node := object
			for part := range strings.SplitSeq(path, ".") {
				node = field(node, part)
			}
			if node == nil {
				return nil, false, nil
			}
			if err := validateYAMLPropertyPrecision(node, path); err != nil {
				return nil, true, err
			}
			var value any
			err := node.Decode(&value)
			return value, true, err
		}
	}
	result.Form, err = propertyTarget(formProperties, lookup(field(engine.root, "form")), extended)
	if err != nil {
		return PropertyGridData{}, err
	}
	for id, control := range engine.controls {
		if field(control.node, "id") == nil {
			continue
		}
		target, err := propertyTarget(editablePropertiesForControl(control.model, extended), lookup(control.node), extended)
		if err != nil {
			return PropertyGridData{}, err
		}
		result.Controls[id] = target
		if strings.EqualFold(control.model.Type, "TabStrip") {
			tabs := field(control.node, "tabs")
			targets := make([]PropertyTarget, 0)
			if tabs != nil && tabs.Kind == yaml.SequenceNode {
				targets = make([]PropertyTarget, 0, len(tabs.Content))
				for _, tab := range tabs.Content {
					target, err := propertyTarget(EditableTabProperties(), lookup(tab))
					if err != nil {
						return PropertyGridData{}, err
					}
					targets = append(targets, target)
				}
			}
			if result.Tabs == nil {
				result.Tabs = map[string][]PropertyTarget{}
			}
			result.Tabs[id] = targets
		}
	}
	return result, nil
}
