package edit

import (
	"cmp"
	"errors"
	"fmt"
	"math"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const geometryTolerance = 0.000001

type controlGeometry struct {
	control       spec.FormSpecControl
	left, top     float64
	width, height float64
	parentWidth   float64
	parentHeight  float64
}

func applyProductivity(input spec.SpecInput, source []byte, operation Operation) (Result, error) {
	model, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return Result{}, validationError(-1, operation, err)
	}
	if err := validateProductivityOperation(operation); err != nil {
		return Result{}, err
	}
	if operation.Type == PasteControls {
		return applyPasteControls(input, source, model, operation)
	}
	expanded, err := expandProductivityOperation(model, operation)
	if err != nil {
		return Result{}, err
	}
	result, err := Apply(input, source, expanded)
	if err != nil {
		if editError, ok := errors.AsType[*Error](err); ok {
			for index := range editError.Diagnostics {
				editError.Diagnostics[index].OperationIndex = 0
				editError.Diagnostics[index].Operation = operation.Type
			}
		}
		return Result{}, err
	}
	return result, nil
}

func validateProductivityOperation(operation Operation) error {
	invalid := func(message string) error {
		return failure(-1, operation, "UFE003", message, "Supply a complete payload for one productivity operation.")
	}
	if operation.ControlID != "" || operation.Field != "" || operation.Value != nil || operation.ValuePresent ||
		operation.Left != nil || operation.Top != nil || operation.Width != nil || operation.Height != nil ||
		operation.Control != nil || operation.Tab != nil || operation.Index != nil || operation.ToIndex != nil {
		return invalid("Unexpected payload field for productivity operation.")
	}
	finite := func(value float64) bool { return !math.IsNaN(value) && !math.IsInf(value, 0) }
	uniqueIDs := func(ids []string, minimum int) bool {
		if len(ids) < minimum {
			return false
		}
		seen := make(map[string]struct{}, len(ids))
		for _, id := range ids {
			if id == "" {
				return false
			}
			if _, exists := seen[id]; exists {
				return false
			}
			seen[id] = struct{}{}
		}
		return true
	}
	noSelectionFields := func() bool {
		return len(operation.ControlIDs) == 0 && operation.AnchorID == "" && operation.Alignment == "" &&
			operation.Axis == "" && operation.Dimension == "" && operation.Command == "" && operation.ParentID == ""
	}

	switch operation.Type {
	case MoveControls:
		if !noSelectionFields() || operation.Cascade || operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) < 2 {
			return invalid("moveControls requires at least two moves and no unrelated fields.")
		}
		seen := make(map[string]struct{}, len(operation.Moves))
		for _, move := range operation.Moves {
			if move.ControlID == "" || !finite(move.Left) || !finite(move.Top) {
				return invalid("Each move requires a controlId and finite left/top coordinates.")
			}
			if _, exists := seen[move.ControlID]; exists {
				return invalid("moveControls contains a duplicate controlId.")
			}
			seen[move.ControlID] = struct{}{}
		}
	case AlignControls:
		if operation.Cascade || operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) > 0 ||
			operation.Axis != "" || operation.Dimension != "" || operation.Command != "" || operation.ParentID != "" ||
			!uniqueIDs(operation.ControlIDs, 2) || operation.AnchorID == "" {
			return invalid("alignControls requires distinct controlIds and an anchorId.")
		}
		if !slices.Contains([]string{"left", "center", "right", "top", "middle", "bottom"}, operation.Alignment) {
			return invalid("alignment must be left, center, right, top, middle, or bottom.")
		}
	case DistributeControls:
		if operation.Cascade || operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) > 0 ||
			operation.AnchorID != "" || operation.Alignment != "" || operation.Dimension != "" || operation.Command != "" || operation.ParentID != "" ||
			!uniqueIDs(operation.ControlIDs, 3) || !slices.Contains([]string{"horizontal", "vertical"}, operation.Axis) {
			return invalid("distributeControls requires at least three controlIds and a valid axis.")
		}
	case SameSizeControls:
		if operation.Cascade || operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) > 0 ||
			operation.Alignment != "" || operation.Axis != "" || operation.Command != "" || operation.ParentID != "" ||
			!uniqueIDs(operation.ControlIDs, 2) || operation.AnchorID == "" ||
			!slices.Contains([]string{"width", "height", "size"}, operation.Dimension) {
			return invalid("sameSizeControls requires distinct controlIds, an anchorId, and a valid dimension.")
		}
	case RemoveControls:
		if operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) > 0 ||
			operation.AnchorID != "" || operation.Alignment != "" || operation.Axis != "" || operation.Dimension != "" || operation.Command != "" || operation.ParentID != "" ||
			!uniqueIDs(operation.ControlIDs, 1) {
			return invalid("removeControls requires distinct controlIds.")
		}
	case ReorderControls:
		if operation.Cascade || operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) > 0 ||
			operation.AnchorID != "" || operation.Alignment != "" || operation.Axis != "" || operation.Dimension != "" || operation.ParentID != "" ||
			!uniqueIDs(operation.ControlIDs, 1) || !slices.Contains([]string{"front", "forward", "backward", "back"}, operation.Command) {
			return invalid("reorderControls requires distinct controlIds and a valid command.")
		}
	case SetTabOrder:
		if operation.Cascade || operation.Snapshot != nil || len(operation.Allocations) > 0 || operation.Offset != nil || len(operation.Moves) > 0 ||
			operation.AnchorID != "" || operation.Alignment != "" || operation.Axis != "" || operation.Dimension != "" || operation.Command != "" ||
			len(operation.ControlIDs) > 0 && !uniqueIDs(operation.ControlIDs, 1) {
			return invalid("setTabOrder requires a complete ordered controlIds list.")
		}
	case PasteControls:
		if operation.Cascade || len(operation.Moves) > 0 || len(operation.ControlIDs) > 0 || operation.AnchorID != "" || operation.Alignment != "" || operation.Axis != "" || operation.Dimension != "" || operation.Command != "" ||
			operation.Snapshot == nil || len(operation.Allocations) == 0 || operation.Offset == nil ||
			!finite(operation.Offset.Left) || !finite(operation.Offset.Top) {
			return invalid("pasteControls requires snapshot, allocations, parentId, and finite offsets.")
		}
	default:
		return invalid("Unknown productivity operation.")
	}
	return nil
}

func expandProductivityOperation(model spec.FormSpec, operation Operation) ([]Operation, error) {
	switch operation.Type {
	case MoveControls:
		return expandMoveControls(model, operation)
	case AlignControls:
		return expandAlignControls(model, operation)
	case DistributeControls:
		return expandDistributeControls(model, operation)
	case SameSizeControls:
		return expandSameSizeControls(model, operation)
	case RemoveControls:
		if _, err := selectedGeometries(model, operation.ControlIDs, operation); err != nil {
			return nil, err
		}
		result := make([]Operation, 0, len(operation.ControlIDs))
		for _, id := range operation.ControlIDs {
			result = append(result, Operation{Type: RemoveControl, ControlID: id, Cascade: operation.Cascade})
		}
		return result, nil
	case ReorderControls:
		return expandReorderControls(model, operation)
	case SetTabOrder:
		return expandTabOrder(model, operation)
	default:
		return nil, failure(-1, operation, "UFE003", "Unknown productivity operation.", "Use a supported productivity operation.")
	}
}

func expandMoveControls(model spec.FormSpec, operation Operation) ([]Operation, error) {
	ids := make([]string, len(operation.Moves))
	for index, move := range operation.Moves {
		ids[index] = move.ControlID
	}
	states, err := selectedGeometries(model, ids, operation)
	if err != nil {
		return nil, err
	}
	deltaLeft := stableGeometryDelta(operation.Moves[0].Left - states[0].left)
	deltaTop := stableGeometryDelta(operation.Moves[0].Top - states[0].top)
	for index, move := range operation.Moves[1:] {
		state := states[index+1]
		if !sameGeometryValue(stableGeometryDelta(move.Left-state.left), deltaLeft) || !sameGeometryValue(stableGeometryDelta(move.Top-state.top), deltaTop) {
			return nil, failure(-1, operation, "UFE003", "moveControls must apply one shared translation to every selected control.", "Supply a common parent-relative delta for all moves.")
		}
	}
	minLeftDelta, maxLeftDelta := math.Inf(-1), math.Inf(1)
	minTopDelta, maxTopDelta := math.Inf(-1), math.Inf(1)
	for _, state := range states {
		if !validControlSize(state) || state.width > state.parentWidth || state.height > state.parentHeight {
			return nil, failure(-1, operation, "UFE003", "A selected control cannot fit within its parent bounds.", "Resize or reposition the controls before moving the group.")
		}
		if state.left < 0 || state.top < 0 || state.left+state.width > state.parentWidth || state.top+state.height > state.parentHeight {
			return nil, failure(-1, operation, "UFE003", "A selected control is already outside its parent bounds.", "Correct the invalid source geometry before moving the group.")
		}
		minLeftDelta = math.Max(minLeftDelta, -state.left)
		maxLeftDelta = math.Min(maxLeftDelta, state.parentWidth-state.left-state.width)
		minTopDelta = math.Max(minTopDelta, -state.top)
		maxTopDelta = math.Min(maxTopDelta, state.parentHeight-state.top-state.height)
	}
	if minLeftDelta > maxLeftDelta || minTopDelta > maxTopDelta {
		return nil, failure(-1, operation, "UFE003", "The selected group cannot fit within its parent bounds.", "Resize the controls or choose a parent with more space.")
	}
	leftDelta := max(minLeftDelta, min(deltaLeft, maxLeftDelta))
	topDelta := max(minTopDelta, min(deltaTop, maxTopDelta))
	for index := range states {
		state := &states[index]
		if leftDelta != 0 {
			state.left += leftDelta
		}
		if topDelta != 0 {
			state.top += topDelta
		}
	}
	return geometryOperations(states, operation)
}

func expandAlignControls(model spec.FormSpec, operation Operation) ([]Operation, error) {
	states, err := selectedGeometries(model, operation.ControlIDs, operation)
	if err != nil {
		return nil, err
	}
	if !slices.Contains(operation.ControlIDs, operation.AnchorID) {
		return nil, failure(-1, operation, "UFE003", "The alignment anchor must be selected.", "Include anchorId in controlIds.")
	}
	anchor := states[slices.IndexFunc(states, func(state controlGeometry) bool { return state.control.ID == operation.AnchorID })]
	for index := range states {
		state := &states[index]
		if state.control.ID == operation.AnchorID {
			if err := validateControlGeometry(model, *state); err != nil {
				return nil, failure(-1, operation, "UFE003", err.Error(), "Correct the anchor geometry before aligning the selection.")
			}
			continue
		}
		oldLeft, oldTop := state.left, state.top
		switch operation.Alignment {
		case "left":
			state.left = anchor.left
		case "center":
			state.left = anchor.left + (anchor.width-state.width)/2
		case "right":
			state.left = anchor.left + anchor.width - state.width
		case "top":
			state.top = anchor.top
		case "middle":
			state.top = anchor.top + (anchor.height-state.height)/2
		case "bottom":
			state.top = anchor.top + anchor.height - state.height
		}
		state.left = oldLeft + stableGeometryDelta(state.left-oldLeft)
		state.top = oldTop + stableGeometryDelta(state.top-oldTop)
		if err := validateControlGeometry(model, *state); err != nil {
			return nil, failure(-1, operation, "UFE003", err.Error(), "Choose an alignment whose result fits every parent boundary.")
		}
	}
	return geometryOperations(states, operation)
}

func expandDistributeControls(model spec.FormSpec, operation Operation) ([]Operation, error) {
	states, err := selectedGeometries(model, operation.ControlIDs, operation)
	if err != nil {
		return nil, err
	}
	parentID := states[0].control.ParentID
	siblings := orderedChildren(model.Controls, parentID, "")
	selected := make(map[string]bool, len(operation.ControlIDs))
	for _, id := range operation.ControlIDs {
		selected[id] = true
	}
	ordered := make([]controlGeometry, 0, len(states))
	for _, sibling := range siblings {
		if selected[sibling.ID] {
			ordered = append(ordered, geometryForControl(model, sibling))
		}
	}
	axis := func(state controlGeometry) (position, size, limit float64) {
		if operation.Axis == "horizontal" {
			return state.left, state.width, state.parentWidth
		}
		return state.top, state.height, state.parentHeight
	}
	slices.SortStableFunc(ordered, func(left, right controlGeometry) int {
		leftPosition, _, _ := axis(left)
		rightPosition, _, _ := axis(right)
		return cmp.Compare(leftPosition, rightPosition)
	})
	firstPosition, _, _ := axis(ordered[0])
	lastPosition, lastSize, _ := axis(ordered[len(ordered)-1])
	totalSize := 0.0
	for _, state := range ordered {
		_, size, limit := axis(state)
		if !validControlSize(state) || size > limit {
			return nil, failure(-1, operation, "UFE003", "A selected control cannot fit within its parent bounds.", "Resize or reposition the controls before distributing them.")
		}
		totalSize += size
	}
	gap := (lastPosition + lastSize - firstPosition - totalSize) / float64(len(ordered)-1)
	if gap < -geometryTolerance {
		return nil, failure(-1, operation, "UFE003", "The selected controls do not have enough space for a non-overlapping distribution.", "Choose a larger parent or reduce control sizes.")
	}
	gap = max(0, gap)
	precedingSize := 0.0
	for index := range ordered {
		state := &ordered[index]
		if index > 0 && index < len(ordered)-1 {
			position := firstPosition + precedingSize + gap*float64(index)
			currentPosition, _, _ := axis(*state)
			displacement := stableGeometryDelta(position - currentPosition)
			if operation.Axis == "horizontal" {
				state.left += displacement
			} else {
				state.top += displacement
			}
		}
		if err := validateControlGeometry(model, *state); err != nil {
			return nil, failure(-1, operation, "UFE003", err.Error(), "Choose a distribution whose result fits every parent boundary.")
		}
		_, size, _ := axis(*state)
		precedingSize += size
	}
	return geometryOperations(ordered, operation)
}

func expandSameSizeControls(model spec.FormSpec, operation Operation) ([]Operation, error) {
	states, err := selectedGeometries(model, operation.ControlIDs, operation)
	if err != nil {
		return nil, err
	}
	if !slices.Contains(operation.ControlIDs, operation.AnchorID) {
		return nil, failure(-1, operation, "UFE003", "The size anchor must be selected.", "Include anchorId in controlIds.")
	}
	anchor := states[slices.IndexFunc(states, func(state controlGeometry) bool { return state.control.ID == operation.AnchorID })]
	for index := range states {
		state := &states[index]
		if operation.Dimension == "width" || operation.Dimension == "size" {
			state.width = anchor.width
		}
		if operation.Dimension == "height" || operation.Dimension == "size" {
			state.height = anchor.height
		}
		if err := validateControlGeometry(model, *state); err != nil {
			return nil, failure(-1, operation, "UFE003", err.Error(), "Choose dimensions that fit every selected control's parent.")
		}
	}
	return geometryOperations(states, operation)
}

func expandReorderControls(model spec.FormSpec, operation Operation) ([]Operation, error) {
	states, err := selectedGeometries(model, operation.ControlIDs, operation)
	if err != nil {
		return nil, err
	}
	parentID := states[0].control.ParentID
	current := orderedChildren(model.Controls, parentID, "")
	selected := make(map[string]bool, len(operation.ControlIDs))
	for _, id := range operation.ControlIDs {
		selected[id] = true
	}
	final := slices.Clone(current)
	switch operation.Command {
	case "front":
		final = make([]spec.FormSpecControl, 0, len(current))
		for _, control := range current {
			if !selected[control.ID] {
				final = append(final, control)
			}
		}
		for _, control := range current {
			if selected[control.ID] {
				final = append(final, control)
			}
		}
	case "back":
		final = make([]spec.FormSpecControl, 0, len(current))
		for _, control := range current {
			if selected[control.ID] {
				final = append(final, control)
			}
		}
		for _, control := range current {
			if !selected[control.ID] {
				final = append(final, control)
			}
		}
	case "forward":
		for index := len(final) - 2; index >= 0; index-- {
			if selected[final[index].ID] && !selected[final[index+1].ID] {
				final[index], final[index+1] = final[index+1], final[index]
			}
		}
	case "backward":
		for index := 1; index < len(final); index++ {
			if selected[final[index].ID] && !selected[final[index-1].ID] {
				final[index], final[index-1] = final[index-1], final[index]
			}
		}
	}
	working := slices.Clone(current)
	result := make([]Operation, 0, len(current))
	for target, wanted := range final {
		index := slices.IndexFunc(working, func(control spec.FormSpecControl) bool { return control.ID == wanted.ID })
		if index == target {
			continue
		}
		result = append(result, Operation{Type: ReorderControl, ControlID: wanted.ID, Index: new(target)})
		control := working[index]
		working = slices.Delete(working, index, index+1)
		working = slices.Insert(working, target, control)
	}
	return result, nil
}

func expandTabOrder(model spec.FormSpec, operation Operation) ([]Operation, error) {
	if err := validateProductivityParent(model, operation.ParentID, operation); err != nil {
		return nil, err
	}
	eligible := make(map[string]bool)
	for _, control := range model.Controls {
		if control.ParentID != operation.ParentID || strings.EqualFold(control.Type, "Page") {
			continue
		}
		if _, ok := EditableControlProperties(control.Type)["tabIndex"]; ok {
			eligible[control.ID] = true
		}
	}
	if len(eligible) != len(operation.ControlIDs) {
		return nil, failure(-1, operation, "UFE003", "setTabOrder must contain every eligible immediate child exactly once.", "Refresh and submit the complete parent-local tab order.")
	}
	result := make([]Operation, 0, len(eligible))
	for index, id := range operation.ControlIDs {
		if !eligible[id] {
			return nil, failure(-1, operation, "UFE003", "setTabOrder contains a control outside the eligible immediate-child set.", "Use each eligible child ID once.")
		}
		result = append(result, Operation{Type: SetControlProperty, ControlID: id, Field: "tabIndex", Value: index, ValuePresent: true})
	}
	return result, nil
}

func selectedGeometries(model spec.FormSpec, ids []string, operation Operation) ([]controlGeometry, error) {
	byID := make(map[string]spec.FormSpecControl, len(model.Controls))
	for _, control := range model.Controls {
		byID[control.ID] = control
	}
	states := make([]controlGeometry, 0, len(ids))
	parentID := ""
	for index, id := range ids {
		control, ok := byID[id]
		if !ok {
			return nil, failure(-1, operation, "UFE004", "A selected control ID was not found.", "Use IDs from the current FormSpec.")
		}
		if strings.EqualFold(control.Type, "Page") || !isBuiltInControl(control) {
			return nil, failure(-1, operation, "UFE005", "Page and custom controls cannot be direct productivity targets.", "Select ordinary built-in controls.")
		}
		if index == 0 {
			parentID = control.ParentID
		} else if control.ParentID != parentID {
			return nil, failure(-1, operation, "UFE005", "A layout selection must contain immediate siblings with the same parent.", "Select controls under one parent.")
		}
		states = append(states, geometryForControl(model, control))
	}
	if err := validateProductivityParent(model, parentID, operation); err != nil {
		return nil, err
	}
	return states, nil
}

func validateProductivityParent(model spec.FormSpec, parentID string, operation Operation) error {
	if parentID == "" {
		return nil
	}
	parent, ok := controlByID(model, parentID)
	if !ok {
		return failure(-1, operation, "UFE004", "Productivity parent was not found.", "Choose a current Form, Frame, or Page parent.")
	}
	if !isBuiltInControl(parent) || (!strings.EqualFold(parent.Type, "Frame") && !strings.EqualFold(parent.Type, "Page")) {
		return failure(-1, operation, "UFE005", "Designer productivity requires a Form, Frame, or Page parent.", "Select immediate children of the Form, a built-in Frame, or a built-in Page.")
	}
	return nil
}

func geometryForControl(model spec.FormSpec, control spec.FormSpecControl) controlGeometry {
	observed := control.Observed
	width, height := designerControlDefaults(control.Type)
	parentWidth, parentHeight := geometryParentBounds(model, control.ParentID)
	return controlGeometry{
		control:      control,
		left:         firstNumber(control.Left, observedNumber(observed, "left")),
		top:          firstNumber(control.Top, observedNumber(observed, "top")),
		width:        firstNumber(control.Width, observedNumber(observed, "width"), new(width)),
		height:       firstNumber(control.Height, observedNumber(observed, "height"), new(height)),
		parentWidth:  parentWidth,
		parentHeight: parentHeight,
	}
}

func geometryParentBounds(model spec.FormSpec, parentID string) (float64, float64) {
	if parentID == "" {
		return designerFormBounds(model)
	}
	parent, ok := controlByID(model, parentID)
	if !ok {
		return 0, 0
	}
	if strings.EqualFold(parent.Type, "Page") {
		owner, found := controlByID(model, parent.ParentID)
		if !found {
			return 0, 0
		}
		return contentBounds(geometryForControl(model, owner))
	}
	return contentBounds(geometryForControl(model, parent))
}

func designerFormBounds(model spec.FormSpec) (float64, float64) {
	width := firstNumber(model.Form.Width, new(240.0))
	height := firstNumber(model.Form.Height, new(180.0))
	return max(0, width), max(0, height)
}

func contentBounds(parent controlGeometry) (float64, float64) {
	left, top, right, bottom := 0.0, 0.0, 0.0, 0.0
	switch strings.ToLower(parent.control.Type) {
	case "frame":
		left, top, right, bottom = 0.75, 0.75, 0.75, 0.75
	case "multipage":
		left, top, right, bottom = 1.5, 18, 1.5, 1.5
	}
	return max(0, parent.width-left-right), max(0, parent.height-top-bottom)
}

func firstNumber(values ...*float64) float64 {
	for _, value := range values {
		if value != nil {
			return *value
		}
	}
	return 0
}

func observedNumber(observed *spec.FormSpecObservedControl, field string) *float64 {
	if observed == nil {
		return nil
	}
	switch field {
	case "left":
		return observed.Left
	case "top":
		return observed.Top
	case "width":
		return observed.Width
	case "height":
		return observed.Height
	default:
		return nil
	}
}

func designerControlDefaults(typeName string) (float64, float64) {
	switch strings.ToLower(typeName) {
	case "label":
		return 72, 18
	case "textbox", "combobox":
		return 120, 18
	case "listbox":
		return 120, 72
	case "commandbutton":
		return 72, 24
	case "checkbox", "optionbutton":
		return 72, 18
	case "togglebutton":
		return 72, 18
	case "spinbutton":
		return 18, 36
	case "scrollbar":
		return 120, 18
	case "image":
		return 72, 72
	case "frame":
		return 144, 108
	case "multipage":
		return 240, 180
	case "tabstrip":
		return 240, 48
	case "page":
		return 0, 0
	default:
		return 72, 24
	}
}

func geometryOperations(states []controlGeometry, operation Operation) ([]Operation, error) {
	result := make([]Operation, 0, len(states)*2)
	for _, state := range states {
		current := geometryForControlFromModel(state.control)
		if current.left != state.left {
			result = append(result, Operation{Type: SetControlProperty, ControlID: state.control.ID, Field: "left", Value: state.left, ValuePresent: true})
		}
		if current.top != state.top {
			result = append(result, Operation{Type: SetControlProperty, ControlID: state.control.ID, Field: "top", Value: state.top, ValuePresent: true})
		}
		if current.width != state.width {
			result = append(result, Operation{Type: SetControlProperty, ControlID: state.control.ID, Field: "width", Value: state.width, ValuePresent: true})
		}
		if current.height != state.height {
			result = append(result, Operation{Type: SetControlProperty, ControlID: state.control.ID, Field: "height", Value: state.height, ValuePresent: true})
		}
	}
	return result, nil
}

func geometryForControlFromModel(control spec.FormSpecControl) controlGeometry {
	width, height := designerControlDefaults(control.Type)
	observed := control.Observed
	return controlGeometry{
		control: control,
		left:    firstNumber(control.Left, observedNumber(observed, "left")),
		top:     firstNumber(control.Top, observedNumber(observed, "top")),
		width:   firstNumber(control.Width, observedNumber(observed, "width"), new(width)),
		height:  firstNumber(control.Height, observedNumber(observed, "height"), new(height)),
	}
}

func validControlSize(state controlGeometry) bool {
	return finiteGeometry(state.width) && finiteGeometry(state.height) && state.width >= 1 && state.height >= 1
}

func validateControlGeometry(model spec.FormSpec, state controlGeometry) error {
	requiredWidth, requiredHeight := requiredControlSize(model, state.control.ID)
	if !validControlSize(state) || state.width+geometryTolerance < requiredWidth || state.height+geometryTolerance < requiredHeight {
		return fmt.Errorf("the resulting control size is below the Designer minimum")
	}
	if !finiteGeometry(state.left) || !finiteGeometry(state.top) || state.left < 0 || state.top < 0 ||
		state.width > state.parentWidth || state.height > state.parentHeight ||
		state.left+state.width > state.parentWidth+geometryTolerance || state.top+state.height > state.parentHeight+geometryTolerance {
		return fmt.Errorf("the resulting geometry is outside its parent content bounds")
	}
	return nil
}

func requiredControlSize(model spec.FormSpec, controlID string) (float64, float64) {
	control, ok := controlByID(model, controlID)
	if !ok || (!strings.EqualFold(control.Type, "Frame") && !strings.EqualFold(control.Type, "MultiPage")) {
		return 1, 1
	}
	owners := map[string]bool{controlID: true}
	leftInset, rightInset, topInset, bottomInset := 0.0, 0.0, 0.0, 0.0
	if strings.EqualFold(control.Type, "Frame") {
		leftInset, rightInset, topInset, bottomInset = 0.75, 0.75, 0.75, 0.75
	} else {
		leftInset, rightInset, topInset, bottomInset = 1.5, 1.5, 18, 1.5
		for _, page := range model.Controls {
			if page.ParentID == controlID && strings.EqualFold(page.Type, "Page") {
				owners[page.ID] = true
			}
		}
	}
	width, height := 1.0, 1.0
	for _, child := range model.Controls {
		if !owners[child.ParentID] || strings.EqualFold(child.Type, "Page") {
			continue
		}
		geometry := geometryForControl(model, child)
		width = max(width, geometry.left+geometry.width+leftInset+rightInset)
		height = max(height, geometry.top+geometry.height+topInset+bottomInset)
	}
	return width, height
}

func finiteGeometry(value float64) bool {
	return !math.IsNaN(value) && !math.IsInf(value, 0)
}

func stableGeometryDelta(value float64) float64 {
	return math.Round(value*1e6) / 1e6
}

func sameGeometryValue(left, right float64) bool {
	return math.Abs(left-right) <= geometryTolerance
}
