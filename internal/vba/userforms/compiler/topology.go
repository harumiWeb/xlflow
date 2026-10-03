package compiler

import (
	"cmp"
	"errors"
	"math"
	"reflect"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// CompileEdits validates a complete before snapshot, applies property edits to
// retained controls, then compiles the requested Frame/common-control topology.
// Both stages operate on independent signed models.
func CompileEdits(base *oforms.Form, before, after spec.FormSpec, codePage uint16) (*oforms.Form, error) {
	if _, err := oforms.SerializeForm(base, codePage); err != nil {
		return nil, err
	}
	for _, input := range []struct{ name, coordinates string }{
		{"before", before.CoordinateSystem}, {"after", after.CoordinateSystem},
	} {
		switch input.coordinates {
		case "", "points", "parent-relative":
		default:
			return nil, templateError(base, Invalid, "", input.name+".coordinateSystem", "geometry requires points or parent-relative coordinates")
		}
	}
	old, err := normalizedCopy(before)
	if err != nil {
		return nil, templateError(base, Invalid, "", "before", err.Error())
	}
	candidate, err := cloneFormSpec(after)
	if err != nil {
		return nil, templateError(base, Invalid, "", "after", err.Error())
	}
	oldSelections := map[string]spec.FormSpecControl{}
	for _, control := range old.Controls {
		oldSelections[control.Name] = control
	}
	var omitUnchangedSelection func([]spec.FormSpecControl)
	omitUnchangedSelection = func(controls []spec.FormSpecControl) {
		for i := range controls {
			control := &controls[i]
			previous, retained := oldSelections[control.Name]
			if retained && strings.EqualFold(previous.Type, control.Type) && (strings.EqualFold(control.Type, "MultiPage") || strings.EqualFold(control.Type, "TabStrip")) && reflect.DeepEqual(previous.SelectedIndex, control.SelectedIndex) {
				control.SelectedIndex = nil
			}
			omitUnchangedSelection(control.Controls)
		}
	}
	omitUnchangedSelection(candidate.Controls)
	next, err := normalizedCopy(candidate)
	if err != nil {
		return nil, templateError(base, Invalid, "", "after", err.Error())
	}
	actual, err := projection.Project(base)
	if err != nil {
		return nil, err
	}
	// Existing topology validation binds the entire supplied before tree to
	// persistence, independently of how many controls the new tree retains.
	check := compilation{form: base, before: old, after: old, actual: actual, controls: map[string]*oforms.Control{}}
	if err := check.topology(); err != nil {
		return nil, err
	}
	if next.Form.Name != base.Name {
		return nil, templateError(base, Unsupported, "", "form.name", "form rename is unsupported")
	}
	oldByName := map[string]spec.FormSpecControl{}
	for _, c := range old.Controls {
		oldByName[c.Name] = c
	}
	propertyAfter, err := cloneFormSpec(old)
	if err != nil {
		return nil, err
	}
	propertyAfter.Form = next.Form
	nextByName := map[string]spec.FormSpecControl{}
	for _, c := range next.Controls {
		nextByName[c.Name] = c
	}
	for i, c := range propertyAfter.Controls {
		if n, ok := nextByName[c.Name]; ok && strings.EqualFold(c.Type, n.Type) {
			if n.ProgID != "" && !strings.EqualFold(n.ProgID, c.ProgID) {
				return nil, templateError(base, Unsupported, c.Name, "progId", "control ProgID change is unsupported")
			}
			n.ID, n.ParentID, n.ZIndex, n.ProgID, n.Type = c.ID, c.ParentID, c.ZIndex, c.ProgID, c.Type
			propertyAfter.Controls[i] = n
			if strings.EqualFold(c.Type, "MultiPage") || strings.EqualFold(c.Type, "TabStrip") {
				propertyAfter.Controls[i].SelectedIndex = c.SelectedIndex
				propertyAfter.Controls[i].Tabs = c.Tabs
			}
		}
	}
	updated, err := compilePropertyEdits(base, old, propertyAfter, codePage)
	if err != nil {
		return nil, err
	}
	retainedTabIndexes := map[string]int16{}
	for _, level := range updated.Levels {
		for _, control := range level.Controls {
			retainedTabIndexes[control.Name] = 0
			if control.TabIndex != nil {
				retainedTabIndexes[control.Name] = *control.TabIndex
			}
		}
	}
	children := map[string][]spec.FormSpecControl{}
	byID := map[string]spec.FormSpecControl{}
	for _, c := range next.Controls {
		byID[c.ID] = c
		children[c.ParentID] = append(children[c.ParentID], c)
	}
	for parent := range children {
		slices.SortStableFunc(children[parent], func(a, b spec.FormSpecControl) int { return cmp.Compare(*a.ZIndex, *b.ZIndex) })
	}
	var desired []oforms.TopologyControl
	var walk func(string, int) error
	walk = func(parentID string, depth int) error {
		siblings := children[parentID]
		items := make([]oforms.TopologyControl, len(siblings))
		usedTabIndexes := map[int16]bool{}
		var defaults []*oforms.ControlDefinition
		// Reserve final sibling values before assigning defaults, including
		// retained controls moved from another parent and explicit property aliases.
		for i, c := range siblings {
			parentName := ""
			if parentID != "" {
				parentName = byID[parentID].Name
			}
			item := oforms.TopologyControl{Name: c.Name, Parent: parentName}
			oldControl, retained := oldByName[c.Name]
			if !retained || !strings.EqualFold(oldControl.Type, c.Type) {
				definition, err := generationControlDefinition(c, 0)
				if err != nil {
					detail := generationErrorForControl(base.Name, c.Name, "controls", err)
					return editGenerationError(detail)
				}
				item.Definition = &definition
				explicit := c.TabIndex != nil
				for key := range c.Properties {
					explicit = explicit || strings.EqualFold(strings.TrimSpace(key), "tabIndex")
				}
				if explicit {
					usedTabIndexes[definition.TabIndex] = true
				} else {
					defaults = append(defaults, item.Definition)
				}
			} else {
				usedTabIndexes[retainedTabIndexes[c.Name]] = true
			}
			items[i] = item
		}
		nextTabIndex := 0
		for _, definition := range defaults {
			for nextTabIndex <= math.MaxInt16 && usedTabIndexes[int16(nextTabIndex)] {
				nextTabIndex++
			}
			if nextTabIndex > math.MaxInt16 {
				return templateError(base, Invalid, definition.Name, "tabIndex", "no unused sibling TabIndex is available")
			}
			definition.TabIndex = int16(nextTabIndex)
			usedTabIndexes[definition.TabIndex] = true
		}
		for i, c := range siblings {
			desired = append(desired, items[i])
			if len(children[c.ID]) > 0 {
				if depth+1 >= oforms.MaxNestingDepth {
					return templateError(base, Invalid, c.Name, "controls", "nesting limit")
				}
				if err := walk(c.ID, depth+1); err != nil {
					return err
				}
			}
		}
		return nil
	}
	if err := walk("", 0); err != nil {
		return nil, err
	}
	if len(desired) != len(next.Controls) {
		return nil, templateError(base, Invalid, "", "controls", "unreachable control hierarchy")
	}
	result, err := oforms.ApplyTopology(updated, desired, codePage)
	if err != nil {
		code := Invalid
		if errors.Is(err, oforms.ErrUnsupportedEdit) {
			code = Unsupported
		}
		return nil, &Error{Code: code, Form: base.Name, Property: "controls", Reason: err.Error(), Cause: err}
	}
	return compileTabIntent(base, result, old, next, codePage)
}

func editGenerationError(err error) error {
	if detail, ok := errors.AsType[*Error](err); ok {
		switch detail.Code {
		case GenerationUnsupported:
			detail.Code = Unsupported
		case GenerationConflict:
			detail.Code = Conflict
		default:
			detail.Code = Invalid
		}
	}
	return err
}

func overlayTemplateTopology(base *oforms.Form, baseline, desired []spec.FormSpecControl) ([]spec.FormSpecControl, error) {
	byName := map[string]spec.FormSpecControl{}
	for _, c := range baseline {
		byName[c.Name] = c
	}
	result := make([]spec.FormSpecControl, 0, len(desired))
	for i, c := range desired {
		old, retained := byName[c.Name]
		if retained && strings.EqualFold(old.Type, c.Type) {
			if c.ProgID != "" && !strings.EqualFold(old.ProgID, c.ProgID) {
				return nil, templateError(base, Unsupported, c.Name, "progId", "control ProgID change is unsupported")
			}
			if err := overlayTemplateControl(base, &old, c, i); err != nil {
				return nil, err
			}
			old.ID, old.ParentID, old.ZIndex = c.ID, c.ParentID, c.ZIndex
			if c.SelectedIndex == nil && (strings.EqualFold(c.Type, "MultiPage") || strings.EqualFold(c.Type, "TabStrip")) {
				old.SelectedIndex = nil
			}
			result = append(result, old)
		} else {
			result = append(result, c)
		}
	}
	return result, nil
}
