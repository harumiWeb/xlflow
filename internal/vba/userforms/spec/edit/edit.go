// Package edit applies semantic FormSpec operations without depending on a
// particular editor or rewriting unrelated source.
package edit

import (
	"bytes"
	"cmp"
	"encoding/json"
	"errors"
	"fmt"
	"math"
	"math/big"
	"reflect"
	"slices"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

type OperationType string

const (
	SetFormProperty    OperationType = "setFormProperty"
	SetControlProperty OperationType = "setControlProperty"
	MoveControl        OperationType = "moveControl"
	ResizeControl      OperationType = "resizeControl"
	AddControl         OperationType = "addControl"
	RemoveControl      OperationType = "removeControl"
	SetParent          OperationType = "setParent"
	ReorderControl     OperationType = "reorderControl"
)

// Operation is a tagged, serializable semantic edit. Geometry is in points,
// relative to the owning parent. Value must be a non-null scalar.
type Operation struct {
	Type      OperationType         `json:"type"`
	ControlID string                `json:"controlId,omitempty"`
	Field     string                `json:"field,omitempty"`
	Value     any                   `json:"value,omitempty"`
	Left      *float64              `json:"left,omitempty"`
	Top       *float64              `json:"top,omitempty"`
	Width     *float64              `json:"width,omitempty"`
	Height    *float64              `json:"height,omitempty"`
	Control   *spec.FormSpecControl `json:"control,omitempty"`
	ParentID  string                `json:"parentId,omitempty"`
	Index     *int                  `json:"index,omitempty"`
	Cascade   bool                  `json:"cascade,omitzero"`
}

// SourceEdit replaces [Start, End) in the original UTF-8 source with Text.
// Result edits are ordered, non-overlapping, and apply to the same snapshot.
type SourceEdit struct {
	Start int    `json:"start"`
	End   int    `json:"end"`
	Text  string `json:"text"`
}

type Result struct {
	Source   []byte
	Document spec.FormSpec
	Edits    []SourceEdit
	Warnings []spec.ValidationIssue
}

type Diagnostic struct {
	OperationIndex int           `json:"operationIndex"`
	Operation      OperationType `json:"operation,omitempty"`
	ControlID      string        `json:"controlId,omitempty"`
	Field          string        `json:"field,omitempty"`
	Code           string        `json:"code"`
	Message        string        `json:"message"`
	Suggestion     string        `json:"suggestion,omitempty"`
}

// Error retains every authoritative validation diagnostic, not just the first.
type Error struct {
	Diagnostics []Diagnostic `json:"diagnostics"`
	cause       error
}

func (e *Error) Error() string {
	if e == nil || len(e.Diagnostics) == 0 {
		return ""
	}
	return e.Diagnostics[0].Message
}
func (e *Error) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.cause
}

// Apply validates the original and final FormSpecs. Intermediate operations may
// temporarily violate FormSpec constraints, enabling atomic structural changes.
// A failure always returns a zero Result; neither source nor payload is mutated.
func Apply(input spec.SpecInput, source []byte, operations []Operation) (Result, error) {
	if input.Format == "json" {
		return applyJSON(input, source, operations)
	}
	if input.Format != "yaml" {
		return Result{}, failure(-1, Operation{}, "UFE001", "Unsupported FormSpec source format.", "Use YAML/YML, or JSON for geometry edits.")
	}
	_, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return Result{}, validationError(-1, Operation{}, err)
	}
	e, err := newEngine(source)
	if err != nil {
		return Result{}, validationError(-1, Operation{}, err)
	}
	affected := map[string]int{}
	for index, op := range operations {
		previous := e.model.Controls
		if err := e.apply(op); err != nil {
			return Result{}, annotate(index, op, err)
		}
		id := op.ControlID
		if op.Control != nil {
			id = op.Control.ID
		}
		if id != "" {
			affected[id] = index
		}
		if op.Type == AddControl || op.Type == RemoveControl || op.Type == SetParent {
			for _, c := range previous {
				if c.ID == id && c.ParentID != "" {
					affected[c.ParentID] = index
				}
			}
			if c, ok := e.controls[id]; ok && c.model.ParentID != "" {
				affected[c.model.ParentID] = index
			}
		}
	}
	after, err := spec.ParseFormSpec(input, e.source)
	last, lastIndex := Operation{}, -1
	if len(operations) > 0 {
		lastIndex = len(operations) - 1
		last = operations[lastIndex]
	}
	if err != nil {
		structured := validationError(lastIndex, last, err)
		for i := range structured.Diagnostics {
			d := &structured.Diagnostics[i]
			id := e.controlIDForField(d.Field)
			if id != "" {
				d.ControlID = id
				if index, ok := affected[id]; ok {
					d.OperationIndex = index
					d.Operation = operations[index].Type
				}
			} else if strings.HasPrefix(d.Field, "form.") {
				d.ControlID = ""
				for index := len(operations) - 1; index >= 0; index-- {
					op := operations[index]
					if op.Type == SetFormProperty && (d.Field == "form."+op.Field || strings.HasPrefix("form."+op.Field, d.Field+".")) {
						d.OperationIndex = index
						d.Operation = op.Type
						break
					}
				}
			}
		}
		return Result{}, structured
	}
	if !sameModel(e.model, after) {
		return Result{}, failure(lastIndex, last, "UFE009", "Generated source does not match the intended FormSpec.", "Keep the source unchanged and report this edit.")
	}
	return Result{Source: bytes.Clone(e.source), Document: after, Edits: e.pieces.edits(source), Warnings: after.ValidationWarnings}, nil
}

func (e *engine) controlIDForField(path string) string {
	n := e.root
	id := ""
	path = strings.ReplaceAll(strings.ReplaceAll(path, "[", "."), "]", "")
	for part := range strings.SplitSeq(path, ".") {
		if n == nil {
			break
		}
		if n.Kind == yaml.MappingNode {
			if value := field(n, "id"); value != nil {
				id = value.Value
			}
			n = field(n, part)
		} else if n.Kind == yaml.SequenceNode {
			index, err := strconv.Atoi(part)
			if err != nil || index < 0 || index >= len(n.Content) {
				break
			}
			n = n.Content[index]
		} else {
			break
		}
	}
	return id
}

func failure(index int, op Operation, code, message, suggestion string) *Error {
	return &Error{Diagnostics: []Diagnostic{{OperationIndex: index, Operation: op.Type, ControlID: operationControlID(op), Field: op.Field, Code: code, Message: message, Suggestion: suggestion}}}
}

func operationControlID(op Operation) string {
	if op.Type == AddControl && op.Control != nil {
		return op.Control.ID
	}
	return op.ControlID
}

func annotate(index int, op Operation, err error) error {
	if editErr, ok := errors.AsType[*Error](err); ok {
		for i := range editErr.Diagnostics {
			d := &editErr.Diagnostics[i]
			d.OperationIndex, d.Operation = index, op.Type
			if d.ControlID == "" {
				d.ControlID = operationControlID(op)
			}
			if d.Field == "" {
				d.Field = op.Field
			}
		}
		return editErr
	}
	return validationError(index, op, err)
}

func validationError(index int, op Operation, err error) *Error {
	result := failure(index, op, "UFE002", err.Error(), "Correct the FormSpec or operation and retry.")
	result.cause = err
	if specErr, ok := errors.AsType[*spec.SpecError](err); ok {
		if len(specErr.Issues) > 0 {
			result.Diagnostics = nil
			for _, issue := range specErr.Issues {
				if issue.Severity != spec.SeverityError {
					continue
				}
				result.Diagnostics = append(result.Diagnostics, Diagnostic{OperationIndex: index, Operation: op.Type, ControlID: operationControlID(op), Field: issue.Field, Code: issue.Code, Message: issue.Message, Suggestion: issue.Suggestion})
			}
		} else {
			result.Diagnostics[0].Code = cmp.Or(specErr.Code, "UFE002")
			result.Diagnostics[0].Field = specErr.Field
			result.Diagnostics[0].Suggestion = specErr.Suggestion
		}
	}
	return result
}

func sameModel(a, b spec.FormSpec) bool {
	a.ValidationWarnings, b.ValidationWarnings = nil, nil
	return reflect.DeepEqual(a, b)
}

type controlRef struct {
	node     *yaml.Node
	sequence *yaml.Node
	parent   string
	model    spec.FormSpecControl
}

type engine struct {
	source   []byte
	root     *yaml.Node
	model    spec.FormSpec
	controls map[string]controlRef
	pieces   pieceTable
}

func newEngine(source []byte) (*engine, error) {
	e := &engine{source: bytes.Clone(source), pieces: pieceTable{{start: 0, end: len(source)}}}
	return e, e.reload()
}

func (e *engine) reload() error {
	var document yaml.Node
	if err := yaml.Unmarshal(e.source, &document); err != nil {
		return err
	}
	if len(document.Content) != 1 {
		return fmt.Errorf("expected one YAML document")
	}
	e.root = document.Content[0]
	var raw spec.FormSpec
	if err := e.root.Decode(&raw); err != nil {
		return err
	}
	e.model = spec.NormalizeFormSpec(raw)
	e.controls = map[string]controlRef{}
	index := 0
	var walk func(*yaml.Node, string)
	walk = func(seq *yaml.Node, parent string) {
		if seq == nil {
			return
		}
		for _, n := range seq.Content {
			if index >= len(e.model.Controls) {
				return
			}
			c := e.model.Controls[index]
			index++
			e.controls[c.ID] = controlRef{node: n, sequence: seq, parent: parent, model: c}
			walk(field(n, "controls"), c.ID)
		}
	}
	walk(field(e.root, "controls"), "")
	return nil
}

func (e *engine) apply(op Operation) error {
	if err := checkPayload(op); err != nil {
		return err
	}
	switch op.Type {
	case SetFormProperty:
		return e.setForm(op)
	case AddControl:
		if op.Control == nil || op.Control.ID == "" || op.Control.Name == "" || op.Control.Type == "" || len(op.Control.Controls) > 0 {
			return failure(-1, op, "UFE003", "AddControl requires an explicit ID, name, type and a single flat control.", "Supply control.id, control.name and control.type.")
		}
		if _, err := json.Marshal(op.Control); err != nil {
			return failure(-1, op, "UFE003", "Control payload is not serializable: "+err.Error(), "Supply serializable FormSpec control fields.")
		}
		var node yaml.Node
		if err := node.Encode(*op.Control); err != nil {
			return err
		}
		return e.appendControl(&node)
	case SetControlProperty, MoveControl, ResizeControl, RemoveControl, SetParent, ReorderControl:
		c, ok := e.controls[op.ControlID]
		if !ok {
			return failure(-1, op, "UFE004", "Control ID was not found.", "Use a control ID from the current FormSpec.")
		}
		if e.unsafeTarget(c.node) {
			return unsupported()
		}
		if field(c.node, "id") == nil {
			return failure(-1, op, "UFE005", "The target has no explicit canonical id field.", "Write the resolved ID as an id field before editing.")
		}
		switch op.Type {
		case SetControlProperty:
			if op.Field == "id" || op.Field == "type" || op.Field == "parentId" || op.Field == "zIndex" {
				return failure(-1, op, "UFE005", "This structural field cannot be set as a property.", "Use a structural operation; ID and type are immutable.")
			}
			contract, ok := spec.ControlProperties(c.model.Type)[op.Field]
			if !ok || !contract.IncludeInAuthoring || !scalarContract(contract) {
				return unsupportedProperty(op)
			}
			return e.setControl(op.ControlID, op.Field, op.Value)
		case MoveControl:
			if op.Left == nil || op.Top == nil {
				return missingGeometry(op)
			}
			if err := e.setControl(op.ControlID, "left", *op.Left); err != nil {
				return err
			}
			return e.setControl(op.ControlID, "top", *op.Top)
		case ResizeControl:
			if op.Width == nil || op.Height == nil {
				return missingGeometry(op)
			}
			if err := e.setControl(op.ControlID, "width", *op.Width); err != nil {
				return err
			}
			return e.setControl(op.ControlID, "height", *op.Height)
		case RemoveControl:
			return e.removeControl(op)
		case SetParent:
			return e.setParent(op)
		case ReorderControl:
			return e.reorder(op)
		}
	}
	return failure(-1, op, "UFE003", "Unknown semantic operation.", "Use a supported operation type.")
}

func scalarContract(c spec.PropertyContract) bool {
	return c.ValueType == spec.ValueTypeString || c.ValueType == spec.ValueTypeNumber || c.ValueType == spec.ValueTypeInteger || c.ValueType == spec.ValueTypeBoolean || c.ValueType == spec.ValueTypeAny
}

func unsupportedProperty(op Operation) error {
	return failure(-1, op, "UFE005", "Property is not an editable scalar authoring field.", "Use a supported scalar field for this control type.")
}
func missingGeometry(op Operation) error {
	return failure(-1, op, "UFE003", "Both geometry axes are required.", "Supply both coordinates or both dimensions.")
}
func unsupported() error {
	return failure(-1, Operation{}, "UFE006", "The target uses anchors, aliases or merge keys.", "Expand the affected YAML mapping before editing.")
}

func checkPayload(op Operation) error {
	allowed := map[string]bool{}
	switch op.Type {
	case SetFormProperty:
		allowed["field"], allowed["value"] = true, true
	case SetControlProperty:
		allowed["controlId"], allowed["field"], allowed["value"] = true, true, true
	case MoveControl:
		allowed["controlId"], allowed["left"], allowed["top"] = true, true, true
	case ResizeControl:
		allowed["controlId"], allowed["width"], allowed["height"] = true, true, true
	case AddControl:
		allowed["control"] = true
	case RemoveControl:
		allowed["controlId"], allowed["cascade"] = true, true
	case SetParent:
		allowed["controlId"], allowed["parentId"], allowed["left"], allowed["top"] = true, true, true, true
	case ReorderControl:
		allowed["controlId"], allowed["index"] = true, true
	default:
		return failure(-1, op, "UFE003", "Unknown semantic operation.", "Use a supported operation type.")
	}
	present := []struct {
		name string
		set  bool
	}{{"controlId", op.ControlID != ""}, {"field", op.Field != ""}, {"value", op.Value != nil}, {"left", op.Left != nil}, {"top", op.Top != nil}, {"width", op.Width != nil}, {"height", op.Height != nil}, {"control", op.Control != nil}, {"parentId", op.ParentID != ""}, {"index", op.Index != nil}, {"cascade", op.Cascade}}
	for _, p := range present {
		if p.set && !allowed[p.name] {
			return failure(-1, op, "UFE003", "Unexpected payload field: "+p.name+".", "Supply only fields belonging to this operation.")
		}
	}
	return nil
}

func field(n *yaml.Node, key string) *yaml.Node {
	if n == nil || n.Kind != yaml.MappingNode {
		return nil
	}
	for i := 0; i+1 < len(n.Content); i += 2 {
		if n.Content[i].Value == key {
			return n.Content[i+1]
		}
	}
	return nil
}

func unsafe(n *yaml.Node) bool {
	if n == nil {
		return false
	}
	if n.Anchor != "" || n.Kind == yaml.AliasNode {
		return true
	}
	if n.Kind == yaml.MappingNode {
		for i := 0; i+1 < len(n.Content); i += 2 {
			if n.Content[i].Value == "<<" {
				return true
			}
		}
	}
	for _, child := range n.Content {
		if unsafe(child) {
			return true
		}
	}
	return false
}

func (e *engine) unsafeTarget(n *yaml.Node) bool {
	for parent := n; parent != nil; parent = e.parent(parent) {
		if parent.Anchor != "" || parent.Kind == yaml.AliasNode || field(parent, "<<") != nil {
			return true
		}
	}
	return false
}

func scalar(value any) (*yaml.Node, error) {
	if value == nil {
		return nil, fmt.Errorf("null is not a scalar edit value")
	}
	if number, ok := value.(json.Number); ok {
		if integer, err := number.Int64(); err == nil {
			return scalar(integer)
		}
		var err error
		value, err = number.Float64()
		if err != nil {
			return nil, err
		}
	}
	if text, ok := value.(string); ok {
		// Node.Encode round-trips through YAML and can lose newline-only
		// strings. Keep the supplied scalar value exact from the start.
		return &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: text}, nil
	}
	// Reject arbitrary Go objects before yaml's encoder can invoke custom
	// marshalers or panic on unsupported kinds. JSON callers use these kinds.
	t := reflect.TypeOf(value)
	switch t.Kind() {
	case reflect.String:
		return &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: reflect.ValueOf(value).String()}, nil
	case reflect.Bool, reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64, reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Float32, reflect.Float64:
	default:
		return nil, fmt.Errorf("edit value must be a string, number or boolean")
	}
	var n yaml.Node
	if err := n.Encode(value); err != nil {
		return nil, err
	}
	if n.Kind != yaml.ScalarNode || n.Tag == "!!null" {
		return nil, fmt.Errorf("edit value must be a non-null scalar")
	}
	return &n, nil
}

func (e *engine) setForm(op Operation) error {
	parts := strings.Split(op.Field, ".")
	var contract spec.PropertyContract
	var ok bool
	if len(parts) == 1 {
		contract, ok = spec.UserFormContract().FormProperties[op.Field]
	} else if len(parts) == 2 && parts[0] == "build" {
		contract, ok = spec.LookupFormBuildProperty(parts[1])
		// Operations name source fields, not case-insensitive metadata aliases.
		ok = ok && hasYAMLField(reflect.TypeFor[spec.FormSpecBuildForm](), parts[1])
	}
	if !ok || (len(parts) == 1 && !contract.IncludeInAuthoring) || !scalarContract(contract) {
		return unsupportedProperty(op)
	}
	form := field(e.root, "form")
	if e.unsafeTarget(form) {
		return unsupported()
	}
	n, err := scalar(op.Value)
	if err != nil {
		return err
	}
	if len(parts) == 1 {
		return e.set(form, parts[0], n)
	}
	build := field(form, "build")
	if build == nil {
		build = &yaml.Node{Kind: yaml.MappingNode, Tag: "!!map", Content: []*yaml.Node{{Kind: yaml.ScalarNode, Tag: "!!str", Value: parts[1]}, n}}
		return e.set(form, "build", build)
	}
	return e.set(build, parts[1], n)
}

func hasYAMLField(t reflect.Type, name string) bool {
	for i := range t.NumField() {
		field := t.Field(i)
		if strings.SplitN(field.Tag.Get("yaml"), ",", 2)[0] == name {
			return true
		}
	}
	return false
}

func (e *engine) setControl(id, key string, value any) error {
	c := e.controls[id]
	n, err := scalar(value)
	if err != nil {
		return err
	}
	old := field(c.node, key)
	if old != nil && equalNode(old, n) {
		return nil
	}
	return e.set(c.node, key, n)
}

func equalNode(a, b *yaml.Node) bool {
	var av, bv any
	if a.Decode(&av) != nil || b.Decode(&bv) != nil {
		return false
	}
	if reflect.DeepEqual(av, bv) {
		return true
	}
	// Numeric scalar spelling (20, 20.0, 2e1) is not an authored change.
	an, bn := rational(av), rational(bv)
	return an != nil && bn != nil && an.Cmp(bn) == 0
}

func rational(value any) *big.Rat {
	switch n := value.(type) {
	case int:
		return new(big.Rat).SetInt64(int64(n))
	case int64:
		return new(big.Rat).SetInt64(n)
	case uint64:
		return new(big.Rat).SetInt(new(big.Int).SetUint64(n))
	case float64:
		return new(big.Rat).SetFloat64(n)
	default:
		return nil
	}
}

func (e *engine) reorder(op Operation) error {
	if op.Index == nil {
		return failure(-1, op, "UFE003", "ReorderControl requires an index.", "Supply a zero-based sibling index.")
	}
	c := e.controls[op.ControlID]
	var siblings []spec.FormSpecControl
	for _, other := range e.model.Controls {
		if other.ParentID == c.model.ParentID {
			siblings = append(siblings, other)
		}
	}
	slices.SortStableFunc(siblings, func(a, b spec.FormSpecControl) int { return cmp.Compare(*a.ZIndex, *b.ZIndex) })
	if *op.Index < 0 || *op.Index >= len(siblings) {
		return failure(-1, op, "UFE003", "Sibling index is out of range.", "Use an index within the sibling collection.")
	}
	old := slices.IndexFunc(siblings, func(c spec.FormSpecControl) bool { return c.ID == op.ControlID })
	if old == *op.Index {
		return nil
	}
	moved := siblings[old]
	siblings = slices.Delete(siblings, old, old+1)
	siblings = slices.Insert(siblings, *op.Index, moved)
	// First try changing only the moved control. Expand to the shortest
	// contiguous range that can fit between unchanged neighbors when needed.
	// Big integers keep capacity checks safe even at the int limits.
	for size := 1; size <= len(siblings); size++ {
		for left := max(0, *op.Index-size+1); left <= min(*op.Index, len(siblings)-size); left++ {
			right := left + size
			low, high := math.MinInt, math.MaxInt
			if left > 0 {
				if *siblings[left-1].ZIndex == math.MaxInt {
					continue
				}
				low = *siblings[left-1].ZIndex + 1
			}
			if right < len(siblings) {
				if *siblings[right].ZIndex == math.MinInt {
					continue
				}
				high = *siblings[right].ZIndex - 1
			}
			capacity := new(big.Int).Sub(big.NewInt(int64(high)), big.NewInt(int64(low)))
			if capacity.Cmp(big.NewInt(int64(size-1))) < 0 {
				continue
			}
			first := low
			if left == 0 {
				first = 0
				if right < len(siblings) {
					first = high - (size - 1)
				}
			}
			for i := range size {
				if err := e.setControl(siblings[left+i].ID, "zIndex", first+i); err != nil {
					return err
				}
			}
			return nil
		}
	}
	return fmt.Errorf("cannot assign sibling zIndex values")
}
