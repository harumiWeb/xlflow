package edit

import (
	"bytes"
	"encoding/json"
	"fmt"
	"math"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

type yamlControlRecord struct {
	node     *yaml.Node
	parentID string
	sequence *yaml.Node
}

// CopyControls copies built-in control roots and their descendants while
// preserving authored YAML/JSON fields in a minimal FormSpec wrapper.
func CopyControls(input spec.SpecInput, source []byte, controlIDs []string) (ClipboardSnapshot, spec.FormSpec, error) {
	format, err := clipboardFormat(input.Format)
	if err != nil {
		return ClipboardSnapshot{}, spec.FormSpec{}, err
	}
	input.Format = format
	document, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return ClipboardSnapshot{}, spec.FormSpec{}, err
	}
	if len(controlIDs) == 0 {
		return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("at least one control root is required")
	}
	rootIDs := make(map[string]bool, len(controlIDs))
	controls := make(map[string]spec.FormSpecControl, len(document.Controls))
	for _, control := range document.Controls {
		controls[control.ID] = control
	}
	rootParent := ""
	for index, id := range controlIDs {
		if id == "" || rootIDs[id] {
			return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("controlIds must contain distinct non-empty IDs")
		}
		control, ok := controls[id]
		if !ok {
			return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("control ID %q was not found", id)
		}
		if index == 0 {
			rootParent = control.ParentID
		} else if control.ParentID != rootParent {
			return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("clipboard roots must have the same immediate parent")
		}
		if strings.EqualFold(control.Type, "Page") || !isBuiltInControl(control) {
			return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("clipboard roots must be ordinary built-in controls other than Page")
		}
		rootIDs[id] = true
	}
	rootOrder := make([]string, 0, len(controlIDs))
	for _, sibling := range orderedChildren(document.Controls, rootParent, "") {
		if rootIDs[sibling.ID] {
			rootOrder = append(rootOrder, sibling.ID)
		}
	}

	included := slices.Clone(rootOrder)
	includedSet := make(map[string]bool, len(rootOrder))
	for _, id := range included {
		includedSet[id] = true
	}
	for changed := true; changed; {
		changed = false
		for _, control := range document.Controls {
			if includedSet[control.ParentID] && !includedSet[control.ID] {
				includedSet[control.ID] = true
				included = append(included, control.ID)
				changed = true
			}
		}
	}
	for _, id := range included {
		if !isBuiltInControl(controls[id]) {
			return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("clipboard subtree contains a custom control")
		}
	}

	root, err := yamlRoot(source)
	if err != nil {
		return ClipboardSnapshot{}, spec.FormSpec{}, err
	}
	records := make(map[string]yamlControlRecord, len(document.Controls))
	indexYAMLControlRecords(field(root, "controls"), "", records)
	for _, id := range included {
		if _, ok := records[id]; !ok {
			return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("control ID %q does not resolve to one authored control record", id)
		}
	}

	// Selected roots always lead the snapshot in selection order. Descendants
	// that are represented as flat parentId records follow in source order;
	// lexically nested descendants remain inside their original parent node.
	topLevel := make([]string, 0, len(included))
	topSeen := make(map[string]bool, len(included))
	for _, id := range rootOrder {
		topLevel = append(topLevel, id)
		topSeen[id] = true
	}
	for _, control := range document.Controls {
		id := control.ID
		if !includedSet[id] || topSeen[id] {
			continue
		}
		ancestor := records[id].parentID
		nested := false
		for ancestor != "" {
			if includedSet[ancestor] {
				nested = true
				break
			}
			ancestor = records[ancestor].parentID
		}
		if !nested {
			topLevel = append(topLevel, id)
			topSeen[id] = true
		}
	}

	copyRoot := cloneYAMLNode(root)
	originalSequence := field(copyRoot, "controls")
	sequence := &yaml.Node{Kind: yaml.SequenceNode, Tag: "!!seq"}
	if originalSequence != nil && originalSequence.Kind == yaml.SequenceNode {
		sequence.Style = originalSequence.Style
	}
	for _, id := range topLevel {
		node := cloneYAMLNode(records[id].node)
		if rootIDs[id] {
			deleteYAMLField(node, "parentId")
		}
		sequence.Content = append(sequence.Content, node)
	}
	setYAMLFieldNode(copyRoot, "controls", sequence)
	if unsafe(copyRoot) {
		return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("copied subtree or retained FormSpec wrapper uses YAML anchors, aliases or merge keys")
	}
	snapshotText, err := marshalClipboardRoot(copyRoot, format)
	if err != nil {
		return ClipboardSnapshot{}, spec.FormSpec{}, err
	}
	snapshot := ClipboardSnapshot{Format: format, Text: string(snapshotText), ControlIDs: rootOrder}
	copyDocument, err := spec.ParseFormSpec(input, snapshotText)
	if err != nil {
		return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("generated clipboard snapshot is invalid: %w", err)
	}
	if len(copyDocument.Controls) != len(includedSet) {
		return ClipboardSnapshot{}, spec.FormSpec{}, fmt.Errorf("generated clipboard snapshot does not contain the complete copied subtree")
	}
	return snapshot, copyDocument, nil
}

func applyPasteControls(input spec.SpecInput, source []byte, target spec.FormSpec, operation Operation) (Result, error) {
	format, err := clipboardFormat(operation.Snapshot.Format)
	if err != nil {
		return Result{}, pasteFailure(operation, err.Error())
	}
	snapshotInput := input
	snapshotInput.Format = format
	snapshotSource := []byte(operation.Snapshot.Text)
	copied, err := spec.ParseFormSpec(snapshotInput, snapshotSource)
	if err != nil {
		return Result{}, pasteFailure(operation, "clipboard snapshot is not a valid FormSpec: "+err.Error())
	}
	if len(operation.Snapshot.ControlIDs) == 0 {
		return Result{}, pasteFailure(operation, "clipboard snapshot must contain at least one root ID")
	}
	copyControls := make(map[string]spec.FormSpecControl, len(copied.Controls))
	for _, control := range copied.Controls {
		copyControls[control.ID] = control
	}
	rootIDs := make(map[string]bool, len(operation.Snapshot.ControlIDs))
	for _, id := range operation.Snapshot.ControlIDs {
		control, ok := copyControls[id]
		if id == "" || !ok || control.ParentID != "" || rootIDs[id] || strings.EqualFold(control.Type, "Page") || !isBuiltInControl(control) {
			return Result{}, pasteFailure(operation, "clipboard root IDs must identify distinct ordinary built-in root controls")
		}
		rootIDs[id] = true
	}
	included := make(map[string]bool, len(copied.Controls))
	for id := range rootIDs {
		included[id] = true
	}
	for changed := true; changed; {
		changed = false
		for _, control := range copied.Controls {
			if included[control.ParentID] && !included[control.ID] {
				included[control.ID] = true
				changed = true
			}
		}
	}
	if len(included) != len(copied.Controls) {
		return Result{}, pasteFailure(operation, "clipboard snapshot contains controls outside its selected roots")
	}
	for _, control := range copied.Controls {
		if !isBuiltInControl(control) {
			return Result{}, pasteFailure(operation, "clipboard subtree contains a custom control")
		}
	}
	if operation.ParentID != "" {
		if err := validateProductivityParent(target, operation.ParentID, operation); err != nil {
			return Result{}, err
		}
	}

	allocations := make(map[string]ControlAllocation, len(operation.Allocations))
	allocatedIDs := make(map[string]bool, len(operation.Allocations))
	allocatedNames := []string{target.Form.Name}
	for _, control := range target.Controls {
		allocatedIDs[control.ID] = true
		allocatedNames = append(allocatedNames, control.Name)
	}
	nameExists := func(name string) bool {
		return slices.ContainsFunc(allocatedNames, func(existing string) bool { return strings.EqualFold(existing, name) })
	}
	for _, allocation := range operation.Allocations {
		if allocation.SourceID == "" || allocation.ID == "" || allocation.Name == "" || allocations[allocation.SourceID].SourceID != "" {
			return Result{}, pasteFailure(operation, "allocations must provide one non-empty sourceId, id and name per copied control")
		}
		if _, ok := copyControls[allocation.SourceID]; !ok || allocatedIDs[allocation.ID] || nameExists(allocation.Name) {
			return Result{}, failure(-1, operation, "UFE008", "A pasted control ID or name collides with the target document.", "Allocate unique IDs and form-wide case-insensitive names.")
		}
		allocations[allocation.SourceID] = allocation
		allocatedIDs[allocation.ID] = true
		allocatedNames = append(allocatedNames, allocation.Name)
	}
	if len(allocations) != len(copied.Controls) {
		return Result{}, pasteFailure(operation, "allocations must cover every control in the copied subtree exactly once")
	}
	for id := range copyControls {
		if allocations[id].SourceID == "" {
			return Result{}, pasteFailure(operation, fmt.Sprintf("allocation for copied control %q is missing", id))
		}
	}

	copyRoot, err := yamlRoot(snapshotSource)
	if err != nil || unsafe(copyRoot) {
		if err == nil {
			err = fmt.Errorf("clipboard snapshot cannot contain YAML aliases or merge keys")
		}
		return Result{}, pasteFailure(operation, err.Error())
	}
	records := make(map[string]yamlControlRecord, len(copied.Controls))
	indexYAMLControlRecords(field(copyRoot, "controls"), "", records)
	for id := range copyControls {
		if _, ok := records[id]; !ok {
			return Result{}, pasteFailure(operation, fmt.Sprintf("copied control %q has no source record", id))
		}
	}
	geometryByID := make(map[string]controlGeometry, len(rootIDs))
	sourceGeometryByID := make(map[string]controlGeometry, len(rootIDs))
	minLeftDelta, maxLeftDelta := math.Inf(-1), math.Inf(1)
	minTopDelta, maxTopDelta := math.Inf(-1), math.Inf(1)
	parentWidth, parentHeight := geometryParentBounds(target, operation.ParentID)
	for id := range rootIDs {
		control := copyControls[id]
		state := geometryForControl(copied, control)
		sourceGeometryByID[id] = state
		if !validControlSize(state) || state.width > parentWidth || state.height > parentHeight {
			return Result{}, pasteFailure(operation, "a copied root cannot fit within the destination parent bounds")
		}
		desiredLeft, desiredTop := state.left+operation.Offset.Left, state.top+operation.Offset.Top
		state.left, state.top = desiredLeft, desiredTop
		state.parentWidth, state.parentHeight = parentWidth, parentHeight
		geometryByID[id] = state
		minLeftDelta = math.Max(minLeftDelta, -desiredLeft)
		maxLeftDelta = math.Min(maxLeftDelta, parentWidth-desiredLeft-state.width)
		minTopDelta = math.Max(minTopDelta, -desiredTop)
		maxTopDelta = math.Min(maxTopDelta, parentHeight-desiredTop-state.height)
	}
	if minLeftDelta > maxLeftDelta || minTopDelta > maxTopDelta {
		return Result{}, pasteFailure(operation, "the copied group does not fit within the destination parent bounds")
	}
	leftCorrection := max(minLeftDelta, min(0.0, maxLeftDelta))
	topCorrection := max(minTopDelta, min(0.0, maxTopDelta))
	for id, state := range geometryByID {
		state.left += leftCorrection
		state.top += topCorrection
		if err := validateControlGeometry(copied, state); err != nil {
			return Result{}, pasteFailure(operation, err.Error())
		}
		geometryByID[id] = state
	}

	maxZ := -1
	for _, sibling := range orderedChildren(target.Controls, operation.ParentID, "") {
		if sibling.ZIndex != nil && *sibling.ZIndex > maxZ {
			maxZ = *sibling.ZIndex
		}
	}
	maxInt := int(^uint(0) >> 1)
	if maxZ > maxInt-len(rootIDs) {
		return Result{}, failure(-1, operation, "UFE003", "The destination z-order has no remaining integer capacity.", "Reorder existing siblings before pasting.")
	}
	rootOrder := make([]string, 0, len(rootIDs))
	for _, sibling := range orderedChildren(copied.Controls, "", "") {
		if rootIDs[sibling.ID] {
			rootOrder = append(rootOrder, sibling.ID)
		}
	}
	if len(rootOrder) != len(rootIDs) {
		return Result{}, pasteFailure(operation, "clipboard roots do not have a stable sibling order")
	}
	rootZ := make(map[string]int, len(rootIDs))
	for index, id := range rootOrder {
		rootZ[id] = maxZ + index + 1
	}

	newRoots := make([]*yaml.Node, 0, len(field(copyRoot, "controls").Content))
	for _, node := range field(copyRoot, "controls").Content {
		clone := cloneYAMLNode(node)
		if err := rewritePastedControl(clone, copyControls, allocations, rootIDs, operation.ParentID, geometryByID, sourceGeometryByID, rootZ); err != nil {
			return Result{}, pasteFailure(operation, err.Error())
		}
		newRoots = append(newRoots, clone)
	}
	return appendPastedControls(input, source, target, operation, newRoots)
}

func appendPastedControls(input spec.SpecInput, source []byte, target spec.FormSpec, operation Operation, controls []*yaml.Node) (Result, error) {
	if input.Format != "json" {
		engine, err := newEngine(source)
		if err != nil {
			return Result{}, failure(-1, operation, "UFE003", err.Error(), "Use a valid YAML FormSpec source.")
		}
		for _, control := range controls {
			if err := engine.appendControl(control); err != nil {
				return Result{}, failure(-1, operation, "UFE003", err.Error(), "The copied controls could not be inserted atomically.")
			}
		}
		result, err := spec.ParseFormSpec(input, engine.source)
		if err != nil {
			return Result{}, validationError(0, operation, err)
		}
		return Result{Source: engine.source, Document: result, Edits: engine.pieces.edits(source), Warnings: result.ValidationWarnings}, nil
	}
	updated := bytes.Clone(source)
	pieces := pieceTable{{start: 0, end: len(source)}}
	for _, control := range controls {
		root, err := parseJSONSource(updated)
		if err != nil {
			return Result{}, failure(-1, operation, "UFE003", err.Error(), "Use a valid JSON FormSpec source.")
		}
		encoded, err := yamlNodeJSON(control)
		if err != nil {
			return Result{}, failure(-1, operation, "UFE003", err.Error(), "The copied control is not JSON serializable.")
		}
		edit, err := jsonAddControlEdit(updated, root, encoded)
		if err != nil {
			return Result{}, annotate(0, operation, err)
		}
		pieces = pieces.replace(edit)
		updated = applySourceEdits(updated, []SourceEdit{edit})
	}
	result, err := spec.ParseFormSpec(input, updated)
	if err != nil {
		return Result{}, validationError(0, operation, err)
	}
	return Result{Source: updated, Document: result, Edits: pieces.edits(source), Warnings: result.ValidationWarnings}, nil
}

func rewritePastedControl(node *yaml.Node, controls map[string]spec.FormSpecControl, allocations map[string]ControlAllocation, roots map[string]bool, parentID string, geometry, sourceGeometry map[string]controlGeometry, zIndexes map[string]int) error {
	if node == nil || node.Kind != yaml.MappingNode {
		return fmt.Errorf("copied control source record must be a mapping")
	}
	oldID, ok := yamlString(field(node, "id"))
	control, exists := controls[oldID]
	allocation, allocated := allocations[oldID]
	if !ok || !exists || !allocated {
		return fmt.Errorf("copied source control does not match an allocation")
	}
	if err := setYAMLField(node, "id", allocation.ID); err != nil {
		return err
	}
	if err := setYAMLField(node, "name", allocation.Name); err != nil {
		return err
	}
	if roots[oldID] {
		if parentID == "" {
			deleteYAMLField(node, "parentId")
		} else if err := setYAMLField(node, "parentId", parentID); err != nil {
			return err
		}
		state := geometry[oldID]
		sourceState := sourceGeometry[oldID]
		if state.left != sourceState.left {
			if err := setYAMLField(node, "left", state.left); err != nil {
				return err
			}
		}
		if state.top != sourceState.top {
			if err := setYAMLField(node, "top", state.top); err != nil {
				return err
			}
		}
		if err := setYAMLField(node, "zIndex", zIndexes[oldID]); err != nil {
			return err
		}
	} else if control.ParentID != "" && field(node, "parentId") != nil {
		parentAllocation, ok := allocations[control.ParentID]
		if !ok {
			return fmt.Errorf("copied descendant references a parent outside the clipboard subtree")
		}
		if err := setYAMLField(node, "parentId", parentAllocation.ID); err != nil {
			return err
		}
	}
	if children := field(node, "controls"); children != nil && children.Kind == yaml.SequenceNode {
		for _, child := range children.Content {
			if err := rewritePastedControl(child, controls, allocations, roots, parentID, geometry, sourceGeometry, zIndexes); err != nil {
				return err
			}
		}
	}
	return nil
}

func yamlRoot(source []byte) (*yaml.Node, error) {
	var document yaml.Node
	if err := yaml.Unmarshal(source, &document); err != nil {
		return nil, err
	}
	if len(document.Content) != 1 || document.Content[0].Kind != yaml.MappingNode {
		return nil, fmt.Errorf("FormSpec source must contain one mapping document")
	}
	return document.Content[0], nil
}

func indexYAMLControlRecords(sequence *yaml.Node, parentID string, records map[string]yamlControlRecord) {
	if sequence == nil || sequence.Kind != yaml.SequenceNode {
		return
	}
	for _, node := range sequence.Content {
		if node.Kind != yaml.MappingNode {
			continue
		}
		id, _ := yamlString(field(node, "id"))
		if id != "" {
			records[id] = yamlControlRecord{node: node, parentID: parentID, sequence: sequence}
		}
		indexYAMLControlRecords(field(node, "controls"), id, records)
	}
}

func cloneYAMLNode(node *yaml.Node) *yaml.Node {
	if node == nil {
		return nil
	}
	clone := *node
	clone.Content = make([]*yaml.Node, len(node.Content))
	for index, child := range node.Content {
		clone.Content[index] = cloneYAMLNode(child)
	}
	return &clone
}

func yamlString(node *yaml.Node) (string, bool) {
	if node == nil || node.Kind != yaml.ScalarNode || node.Tag != "!!str" {
		return "", false
	}
	return node.Value, true
}

func setYAMLFieldNode(mapping *yaml.Node, key string, value *yaml.Node) {
	if mapping == nil || mapping.Kind != yaml.MappingNode {
		return
	}
	for index := 0; index+1 < len(mapping.Content); index += 2 {
		if mapping.Content[index].Value == key {
			mapping.Content[index+1] = value
			return
		}
	}
	mapping.Content = append(mapping.Content, &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: key}, value)
}

func setYAMLField(mapping *yaml.Node, key string, value any) error {
	encoded := &yaml.Node{}
	if err := encoded.Encode(value); err != nil {
		return err
	}
	if old := field(mapping, key); old != nil {
		encoded.HeadComment, encoded.LineComment, encoded.FootComment = old.HeadComment, old.LineComment, old.FootComment
	}
	setYAMLFieldNode(mapping, key, encoded)
	return nil
}

func deleteYAMLField(mapping *yaml.Node, key string) {
	if mapping == nil || mapping.Kind != yaml.MappingNode {
		return
	}
	for index := 0; index+1 < len(mapping.Content); index += 2 {
		if mapping.Content[index].Value == key {
			mapping.Content = slices.Delete(mapping.Content, index, index+2)
			return
		}
	}
}

func marshalClipboardRoot(root *yaml.Node, format string) ([]byte, error) {
	if format == "json" {
		return yamlNodeJSON(root)
	}
	document := &yaml.Node{Kind: yaml.DocumentNode, Content: []*yaml.Node{root}}
	return yaml.Marshal(document)
}

func yamlNodeJSON(node *yaml.Node) ([]byte, error) {
	var result bytes.Buffer
	if err := appendJSONNode(&result, node); err != nil {
		return nil, err
	}
	return result.Bytes(), nil
}

func appendJSONNode(result *bytes.Buffer, node *yaml.Node) error {
	if node == nil {
		return fmt.Errorf("cannot encode an empty YAML node as JSON")
	}
	switch node.Kind {
	case yaml.MappingNode:
		result.WriteByte('{')
		for index := 0; index+1 < len(node.Content); index += 2 {
			if index > 0 {
				result.WriteByte(',')
			}
			key := node.Content[index]
			if key.Kind != yaml.ScalarNode || key.Tag != "!!str" {
				return fmt.Errorf("FormSpec JSON object keys must be strings")
			}
			encodedKey, err := json.Marshal(key.Value)
			if err != nil {
				return err
			}
			result.Write(encodedKey)
			result.WriteByte(':')
			if err := appendJSONNode(result, node.Content[index+1]); err != nil {
				return err
			}
		}
		result.WriteByte('}')
	case yaml.SequenceNode:
		result.WriteByte('[')
		for index, child := range node.Content {
			if index > 0 {
				result.WriteByte(',')
			}
			if err := appendJSONNode(result, child); err != nil {
				return err
			}
		}
		result.WriteByte(']')
	case yaml.ScalarNode:
		switch node.Tag {
		case "!!null":
			result.WriteString("null")
		case "!!bool":
			if node.Value != "true" && node.Value != "false" {
				return fmt.Errorf("invalid boolean scalar %q", node.Value)
			}
			result.WriteString(node.Value)
		case "!!int", "!!float":
			raw := []byte(node.Value)
			if json.Valid(raw) {
				result.Write(raw)
				return nil
			}
			var value any
			if err := node.Decode(&value); err != nil {
				return err
			}
			encoded, err := json.Marshal(value)
			if err != nil {
				return err
			}
			result.Write(encoded)
		case "!!str":
			encoded, err := json.Marshal(node.Value)
			if err != nil {
				return err
			}
			result.Write(encoded)
		default:
			var value any
			if err := node.Decode(&value); err != nil {
				return err
			}
			encoded, err := json.Marshal(value)
			if err != nil {
				return err
			}
			result.Write(encoded)
		}
	default:
		return fmt.Errorf("unsupported YAML node kind %d in JSON output", node.Kind)
	}
	return nil
}

func clipboardFormat(format string) (string, error) {
	switch strings.ToLower(format) {
	case "yaml", "yml":
		return "yaml", nil
	case "json":
		return "json", nil
	default:
		return "", fmt.Errorf("clipboard format must be yaml, yml or json")
	}
}

func isBuiltInControl(control spec.FormSpecControl) bool {
	contract, ok := spec.LookupControlContract(control.Type)
	if !ok {
		return false
	}
	if control.ProgID == "" {
		return true
	}
	progContract, known := spec.LookupControlContractByProgID(control.ProgID)
	return known && strings.EqualFold(progContract.Type, contract.Type)
}

func pasteFailure(operation Operation, message string) error {
	return failure(-1, operation, "UFE003", message, "Refresh the clipboard and destination, then retry.")
}
