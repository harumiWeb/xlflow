package edit

import (
	"encoding/json"
	"fmt"
	"maps"
	"math"
	"reflect"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

// Only these explicitly modeled properties accept editor payloads. In
// particular, the unchecked properties bag is never an object-write API.
func editorPropertyName(name string) bool {
	return slices.Contains([]string{"font", "foreColor", "backColor", "picture", "textAlign", "borderStyle", "pictureAlignment", "pictureSizeMode"}, name)
}

func typedPropertyContract(name string, contract spec.PropertyContract) bool {
	return editorPropertyName(name) && contract.ValueType == spec.ValueTypeObject && contract.Editor != ""
}

func EditableFormEditorProperties() map[string]spec.PropertyContract {
	result := EditableFormProperties()
	for name, contract := range spec.UserFormContract().FormProperties {
		if contract.IncludeInAuthoring && editorPropertyName(name) && hasYAMLField(reflect.TypeFor[spec.FormSpecForm](), name) {
			result[name] = contract
		}
	}
	return result
}

func EditableControlEditorProperties(typeName string) map[string]spec.PropertyContract {
	result := EditableControlProperties(typeName)
	for name, contract := range spec.ControlProperties(typeName) {
		if contract.IncludeInAuthoring && editorPropertyName(name) && hasYAMLField(reflect.TypeFor[spec.FormSpecControl](), name) {
			result[name] = contract
		}
	}
	return result
}

func editablePropertiesForControl(control spec.FormSpecControl, extended bool) map[string]spec.PropertyContract {
	if extended && (control.ProgID == "" || spec.ProgIDMatchesControlType(control.Type, control.ProgID)) {
		return EditableControlEditorProperties(control.Type)
	}
	return EditableControlProperties(control.Type)
}

func propertyDescriptor(name string, contract spec.PropertyContract) PropertyDescriptor {
	d := PropertyDescriptor{Field: name, ValueType: contract.ValueType, Required: contract.Required, Nullable: !contract.Required, AllowedValues: slices.Clone(contract.AllowedValues), Editor: contract.Editor, Minimum: contract.Minimum, Maximum: contract.Maximum, ExclusiveMinimum: contract.ExclusiveMinimum}
	if contract.Editor == "picture" {
		d.Nullable = false
	}
	if d.Editor == "" && contract.ValueType == spec.ValueTypeBoolean {
		d.Editor = "boolean"
	} else if d.Editor == "" && len(d.AllowedValues) > 0 {
		d.Editor = "enum"
	}
	labels := map[string][]string{
		"textAlign":        {"Left", "Center", "Right"},
		"borderStyle":      {"None", "Single"},
		"pictureAlignment": {"TopLeft", "TopRight", "Center", "BottomLeft", "BottomRight"},
		"pictureSizeMode":  {"Clip", "Stretch", "Zoom"},
	}
	if names := labels[name]; len(names) == len(d.AllowedValues) {
		for i, value := range d.AllowedValues {
			d.Suggestions = append(d.Suggestions, PropertySuggestion{Value: value, Label: names[i]})
		}
	}
	for _, name := range slices.Sorted(maps.Keys(contract.Children)) {
		child := propertyDescriptor(name, contract.Children[name])
		if contract.Editor == "font" && name == "charset" {
			child.Suggestions = []PropertySuggestion{{0, "ANSI"}, {1, "Default"}, {2, "Symbol"}, {128, "Shift JIS"}, {129, "Hangul"}, {134, "GB2312"}, {136, "Big5"}, {161, "Greek"}, {162, "Turkish"}, {177, "Hebrew"}, {178, "Arabic"}, {186, "Baltic"}, {204, "Cyrillic"}, {222, "Thai"}, {238, "East Europe"}, {255, "OEM"}}
		}
		if contract.Editor == "color" && name == "value" {
			for index := range 256 {
				color := spec.ColorFromOLE(0x80000000 | uint32(index))
				if color.Kind == "system" {
					child.Suggestions = append(child.Suggestions, PropertySuggestion{Value: color.Value, Label: color.Value})
				}
			}
		}
		d.Children = append(d.Children, child)
	}
	return d
}

func validateYAMLPropertyPrecision(node *yaml.Node, path string) error {
	if node.Kind == yaml.ScalarNode && node.Tag == "!!float" {
		var value any
		if err := node.Decode(&value); err != nil {
			return err
		}
		original := rational(json.Number(strings.ReplaceAll(node.Value, "_", "")))
		encoded, err := json.Marshal(value)
		decoded := rational(json.Number(encoded))
		if err != nil || original == nil || decoded == nil || original.Cmp(decoded) != 0 {
			return propertyPrecisionError(path)
		}
	}
	if node.Kind == yaml.MappingNode {
		for i := 0; i+1 < len(node.Content); i += 2 {
			if err := validateYAMLPropertyPrecision(node.Content[i+1], path+"."+node.Content[i].Value); err != nil {
				return err
			}
		}
	}
	return nil
}

func checkMetadataValue(value any, contract spec.PropertyContract) error {
	if contract.ValueType == spec.ValueTypeObject && contract.Editor != "" {
		return checkTypedValue(Operation{Field: contract.Editor, Value: value, ValuePresent: true}, contract)
	}
	_, err := scalar(value)
	return err
}

func checkTypedValue(op Operation, contract spec.PropertyContract) error {
	if op.Value == nil {
		if contract.Editor == "picture" {
			return failure(-1, op, "UFE003", "Picture null does not remove an image.", "Use picture: {remove: true} to explicitly remove an image.")
		}
		return nil
	}
	object, ok := op.Value.(map[string]any)
	if !ok {
		return failure(-1, op, "UFE003", "Typed property requires an object.", "Use the supplied property editor contract.")
	}
	for name, value := range object {
		child, known := contract.Children[name]
		if !known || !scalarContract(child) {
			return failure(-1, op, "UFE003", "Unknown typed property member: "+name+".", "Use only members described by the property metadata.")
		}
		if value == nil {
			if child.Required {
				return failure(-1, op, "UFE003", "Required typed property member cannot be null: "+name+".", "Supply a value for the required member.")
			}
			continue
		}
		if !propertyNumberRoundTrips(value) {
			return propertyPrecisionError(op.Field + "." + name)
		}
		if _, err := scalar(value); err != nil {
			return err
		}
		valid := true
		switch child.ValueType {
		case spec.ValueTypeString:
			_, valid = value.(string)
		case spec.ValueTypeBoolean:
			_, valid = value.(bool)
		case spec.ValueTypeNumber, spec.ValueTypeInteger:
			n := rational(value)
			valid = n != nil && (child.ValueType != spec.ValueTypeInteger || n.IsInt())
		}
		if !valid {
			return failure(-1, op, "UFE003", "Invalid typed property member: "+name+".", "Use the value type described by the property metadata.")
		}
	}
	for name, child := range contract.Children {
		if _, exists := object[name]; child.Required && !exists {
			return failure(-1, op, "UFE003", "Missing typed property member: "+name+".", "Supply all required members.")
		}
	}
	return nil
}

func propertyNode(value any) (*yaml.Node, error) {
	object, ok := value.(map[string]any)
	if !ok {
		return scalar(value)
	}
	n := &yaml.Node{Kind: yaml.MappingNode, Tag: "!!map"}
	for _, key := range slices.Sorted(maps.Keys(object)) {
		child, err := scalar(object[key])
		if err != nil {
			return nil, err
		}
		n.Content = append(n.Content, &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: key}, child)
	}
	return n, nil
}

func (e *engine) propertyOwner(op Operation) *yaml.Node {
	if op.Type == SetFormProperty || op.ControlID == "" {
		return field(e.root, "form")
	}
	return e.controls[op.ControlID].node
}

func (e *engine) setAuthoredProperty(op Operation, contract spec.PropertyContract) error {
	if !editorPropertyName(op.Field) {
		return e.setControl(op.ControlID, op.Field, op.Value)
	}
	owner := e.propertyOwner(op)
	if e.unsafeTarget(owner) || unsafe(field(owner, op.Field)) || unsafe(field(owner, "properties")) {
		return unsupported()
	}
	if properties := field(owner, "properties"); properties != nil && legacyPropertyName(op.Field) != "" {
		count := 0
		for i := 0; i+1 < len(properties.Content); i += 2 {
			if strings.EqualFold(properties.Content[i].Value, legacyPropertyName(op.Field)) {
				count++
			}
		}
		if count > 1 {
			return failure(-1, op, "UFE008", "Ambiguous legacy property aliases.", "Keep one canonical legacy property key before editing.")
		}
	}
	if op.Field == "font" && field(owner, "font") == nil {
		if properties := field(owner, "properties"); properties != nil {
			for i := 0; i+1 < len(properties.Content); i += 2 {
				if strings.EqualFold(properties.Content[i].Value, "font") {
					var legacy any
					if err := properties.Content[i+1].Decode(&legacy); err != nil {
						return err
					}
					value, err := mergeLegacyFont(op, contract, legacy)
					if err != nil {
						return err
					}
					op.Value = value
				}
			}
		}
	}
	n, err := propertyNode(op.Value)
	if err != nil {
		return err
	}
	old := field(owner, op.Field)
	if old != nil && old.Kind == yaml.MappingNode && n.Kind == yaml.MappingNode && op.Field != "picture" {
		// Font and other typed objects are partial. Reload the owner after
		// each local patch because source positions may have shifted.
		for i := 0; i < len(n.Content); i += 2 {
			if err := e.set(field(e.propertyOwner(op), op.Field), n.Content[i].Value, n.Content[i+1]); err != nil {
				return err
			}
		}
	} else if old != nil && !equalNode(old, n) && (old.Kind != yaml.ScalarNode || n.Kind != yaml.ScalarNode) {
		start, end := e.offset(old), e.end(old)
		text, err := encodeContainer(n, e.newline())
		if err != nil {
			return err
		}
		// Encode replacement objects in flow form, keeping the existing key,
		// indentation, neighboring fields and trailing comments unchanged.
		if n.Kind == yaml.MappingNode {
			n.Style = yaml.FlowStyle
			text, err = encodeContainer(n, e.newline())
			if err != nil {
				return err
			}
		}
		*old = *n
		if err := e.publish(SourceEdit{Start: start, End: end, Text: text}); err != nil {
			return err
		}
	} else if err := e.set(owner, op.Field, n); err != nil {
		return err
	}
	if op.Type == SetControlProperty {
		return e.removeLegacyProperty(op)
	}
	return nil
}

func legacyPropertyName(name string) string {
	if !editorPropertyName(name) || name == "picture" {
		return ""
	}
	return strings.ToUpper(name[:1]) + name[1:]
}

func (e *engine) removeLegacyProperty(op Operation) error {
	if legacyPropertyName(op.Field) == "" {
		return nil
	}
	properties := field(e.propertyOwner(op), "properties")
	if properties == nil {
		return nil
	}
	for i := 0; i+1 < len(properties.Content); i += 2 {
		key, value := properties.Content[i], properties.Content[i+1]
		if !strings.EqualFold(key.Value, legacyPropertyName(op.Field)) {
			continue
		}
		if properties.Style&yaml.FlowStyle != 0 {
			start, end := e.offset(properties), e.end(properties)
			properties.Content = slices.Delete(properties.Content, i, i+2)
			text, err := encodeContainer(properties, e.newline())
			if err != nil {
				return err
			}
			return e.publish(SourceEdit{Start: start, End: end, Text: text})
		}
		start, end := e.lineStart(e.offset(key)), e.lineEnd(e.end(value))
		properties.Content = slices.Delete(properties.Content, i, i+2)
		if len(properties.Content) == 0 {
			// Retain the properties key as an empty mapping; no unrelated bag
			// field or arbitrary property is removed implicitly.
			properties.Style = yaml.FlowStyle
			return e.publish(SourceEdit{Start: start, End: end, Text: strings.Repeat(" ", max(0, key.Column-1)) + "{}" + e.newline()})
		}
		return e.publish(SourceEdit{Start: start, End: end})
	}
	return nil
}

func mergeLegacyFont(op Operation, contract spec.PropertyContract, legacy any) (any, error) {
	update, ok := op.Value.(map[string]any)
	if !ok {
		return op.Value, nil
	}
	previous, ok := legacy.(map[string]any)
	if !ok {
		return nil, unsupportedProperty(op)
	}
	merged := maps.Clone(previous)
	maps.Copy(merged, update)
	op.Value = merged
	if err := checkTypedValue(op, contract); err != nil {
		return nil, err
	}
	return merged, nil
}

func geometryValues(op Operation) (map[string]any, error) {
	values := map[string]any{}
	for _, axis := range []struct {
		name  string
		value *float64
	}{{"left", op.Left}, {"top", op.Top}, {"width", op.Width}, {"height", op.Height}} {
		if axis.value == nil {
			continue
		}
		if math.IsNaN(*axis.value) || math.IsInf(*axis.value, 0) || (op.ControlID == "" && (axis.name == "left" || axis.name == "top")) {
			return nil, failure(-1, op, "UFE003", "Invalid geometry axis.", "Use finite control axes, or form width and height.")
		}
		values[axis.name] = *axis.value
	}
	if len(values) == 0 {
		return nil, failure(-1, op, "UFE003", "At least one geometry axis is required.", "Supply only the axes changed in the Position editor.")
	}
	return values, nil
}

func (e *engine) setGeometry(op Operation) error {
	values, err := geometryValues(op)
	if err != nil {
		return err
	}
	if op.ControlID != "" {
		control, ok := e.controls[op.ControlID]
		if !ok {
			return failure(-1, op, "UFE004", "Control ID was not found.", "Use an explicit control ID from the current FormSpec.")
		}
		if field(control.node, "id") == nil || strings.EqualFold(control.model.Type, "Page") {
			return unsupportedProperty(op)
		}
	}
	for _, key := range slices.Sorted(maps.Keys(values)) {
		if op.ControlID == "" {
			err = e.setForm(Operation{Type: SetFormProperty, Field: key, Value: values[key], ValuePresent: true})
		} else {
			err = e.setControl(op.ControlID, key, values[key])
		}
		if err != nil {
			return err
		}
	}
	return nil
}

func legacyEditorValue(name string, value any) (any, error) {
	if name == "foreColor" || name == "backColor" {
		n := rational(value)
		if n == nil || !n.IsInt() || !n.Num().IsInt64() {
			return nil, fmt.Errorf("legacy color is not an OLE integer")
		}
		number := n.Num().Int64()
		if number < math.MinInt32 || number > math.MaxUint32 {
			return nil, fmt.Errorf("legacy color is outside the OLE range")
		}
		color := spec.ColorFromOLE(uint32(number))
		return map[string]any{"kind": color.Kind, "value": color.Value}, nil
	}
	return value, nil
}

// Keep UseNumber on nested editor values as well as scalar requests.
func marshalEditorValue(value any) ([]byte, error) {
	if _, object := value.(map[string]any); object {
		return json.Marshal(value)
	}
	return marshalJSONScalar(value)
}
