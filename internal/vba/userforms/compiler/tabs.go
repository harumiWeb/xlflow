package compiler

import (
	"fmt"
	"math"
	"reflect"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func compileTabIntent(original, current *oforms.Form, before, after spec.FormSpec, codePage uint16) (*oforms.Form, error) {
	actual, err := projection.Project(original)
	if err != nil {
		return nil, err
	}
	oldByName := map[string]spec.FormSpecControl{}
	actualByName := map[string]spec.FormSpecControl{}
	nextByName := map[string]spec.FormSpecControl{}
	for _, c := range before.Controls {
		oldByName[c.Name] = c
	}
	for _, c := range actual.Controls {
		actualByName[c.Name] = c
	}
	for _, c := range after.Controls {
		nextByName[c.Name] = c
	}
	var edits []oforms.TabEdit
	for _, desired := range after.Controls {
		if !strings.EqualFold(desired.Type, "MultiPage") && !strings.EqualFold(desired.Type, "TabStrip") {
			continue
		}
		owner := findModelControl(current.Controls, desired.Name)
		if owner == nil {
			return nil, templateError(current, Invalid, desired.Name, "controls", "tab owner missing")
		}
		old, retained := oldByName[desired.Name]
		retained = retained && strings.EqualFold(old.Type, desired.Type)
		baseline := actualByName[desired.Name]
		var strip *oforms.TabStrip
		if owner.MultiPage != nil {
			copy := *owner.MultiPage.Hidden.TabStrip
			copy.Tabs = append([]oforms.Tab{}, copy.Tabs...)
			strip = &copy
		} else if owner.TabStrip != nil {
			copy := *owner.TabStrip
			copy.Tabs = append([]oforms.Tab{}, copy.Tabs...)
			strip = &copy
		} else {
			return nil, templateError(current, Invalid, desired.Name, "type", "tab state missing")
		}
		if owner.MultiPage != nil {
			for i, page := range owner.MultiPage.Pages {
				intent := nextByName[page.Name]
				previous := oldByName[page.Name]
				evidence := actualByName[page.Name]
				if !strings.EqualFold(previous.Type, "Page") {
					previous = spec.FormSpecControl{}
					evidence = spec.FormSpecControl{}
				}
				tab := &strip.Tabs[i]
				oldBag, err := propertyMap(previous.Properties)
				if err != nil {
					return nil, err
				}
				nextBag, err := propertyMap(intent.Properties)
				if err != nil {
					return nil, err
				}
				for _, field := range []struct {
					path              string
					old, next, actual any
					apply             func(any)
				}{
					{"caption", pointerValue(previous.Caption), pointerValue(intent.Caption), pointerValue(evidence.Caption), func(v any) { tab.Caption = v.(string) }},
					{"enabled", pointerValue(previous.Enabled), pointerValue(intent.Enabled), pointerValue(evidence.Enabled), func(v any) { tab.Enabled = v.(bool) }},
					{"visible", pointerValue(previous.Visible), pointerValue(intent.Visible), pointerValue(evidence.Visible), func(v any) { tab.Visible = v.(bool) }},
				} {
					if !reflect.DeepEqual(oldBag[field.path], nextBag[field.path]) {
						if !reflect.DeepEqual(field.old, field.next) && !reflect.DeepEqual(field.next, nextBag[field.path]) {
							return nil, templateError(current, Conflict, page.Name, field.path, "conflicting page aliases")
						}
						field.old, field.next = oldBag[field.path], nextBag[field.path]
					}
					if reflect.DeepEqual(field.old, field.next) {
						continue
					}
					if field.next == nil {
						return nil, templateError(current, Unsupported, page.Name, field.path, "property reset is unsupported")
					}
					if field.old != nil && !reflect.DeepEqual(field.old, field.actual) {
						return nil, templateError(current, Stale, page.Name, field.path, "before does not match persisted page tab")
					}
					if field.path == "caption" {
						if _, ok := field.next.(string); !ok {
							return nil, templateError(current, Invalid, page.Name, field.path, "requires string")
						}
					} else if _, ok := field.next.(bool); !ok {
						return nil, templateError(current, Invalid, page.Name, field.path, "requires Boolean")
					}
					field.apply(field.next)
				}
				oldProps, err := propertyMap(previous.Properties)
				if err != nil {
					return nil, err
				}
				props, err := propertyMap(intent.Properties)
				if err != nil {
					return nil, err
				}
				evidenceProps, _ := propertyMap(evidence.Properties)
				if evidence.ControlTipText != nil {
					evidenceProps["controltiptext"] = *evidence.ControlTipText
				}
				if evidence.Accelerator != nil {
					evidenceProps["accelerator"] = *evidence.Accelerator
				}
				for _, key := range []string{"controltiptext", "accelerator"} {
					value, found := props[key]
					oldValue := oldProps[key]
					oldTop, nextTop := previous.ControlTipText, intent.ControlTipText
					if key == "accelerator" {
						oldTop, nextTop = previous.Accelerator, intent.Accelerator
					}
					if !reflect.DeepEqual(oldTop, nextTop) {
						if nextTop == nil {
							return nil, templateError(current, Unsupported, page.Name, key, "property reset is unsupported")
						}
						if found && !reflect.DeepEqual(oldValue, value) && value != *nextTop {
							return nil, templateError(current, Conflict, page.Name, key, "conflicting metadata aliases")
						}
						if oldTop != nil && *oldTop != evidenceProps[key] {
							return nil, templateError(current, Stale, page.Name, key, "before metadata is stale")
						}
						value, found = *nextTop, true
					}
					if reflect.DeepEqual(oldValue, value) {
						continue
					}
					if !found {
						return nil, templateError(current, Unsupported, page.Name, "properties."+key, "property reset is unsupported")
					}
					text, ok := value.(string)
					if !ok {
						return nil, templateError(current, Invalid, page.Name, "properties."+key, "requires string")
					}
					if oldValue != nil && !reflect.DeepEqual(oldValue, evidenceProps[key]) {
						return nil, templateError(current, Stale, page.Name, "properties."+key, "before does not match persisted page tab")
					}
					if key == "accelerator" {
						tab.Accelerator = text
					} else {
						tab.ControlTipText = text
					}
				}
			}
		} else if desired.Tabs != nil && (!retained || !reflect.DeepEqual(old.Tabs, desired.Tabs)) {
			if retained && old.Tabs != nil && !reflect.DeepEqual(old.Tabs, desired.Tabs) && !reflect.DeepEqual(old.Tabs, baseline.Tabs) {
				return nil, templateError(current, Stale, desired.Name, "tabs", "before tabs do not match persisted state")
			}
			strip, err = generationTabs(desired.Tabs, nil, strip)
			if err != nil {
				return nil, editGenerationError(generationErrorForControl(current.Name, desired.Name, "controls", err))
			}
		}
		collectionCount := len(strip.Tabs)
		if owner.MultiPage != nil {
			collectionCount = len(owner.MultiPage.Pages)
		}
		if desired.SelectedIndex != nil && (collectionCount == 0 && *desired.SelectedIndex != -1 || *desired.SelectedIndex < -1 || *desired.SelectedIndex > math.MaxInt32) {
			return nil, templateError(current, Invalid, desired.Name, "selectedIndex", "selection outside final collection")
		}
		if collectionCount > 0 && (!retained || !reflect.DeepEqual(old.SelectedIndex, desired.SelectedIndex)) {
			if desired.SelectedIndex != nil {
				if retained && old.SelectedIndex != nil && !reflect.DeepEqual(old.SelectedIndex, baseline.SelectedIndex) {
					return nil, templateError(current, Stale, desired.Name, "selectedIndex", "before selection is stale")
				}
				strip.SelectedIndex = int32(*desired.SelectedIndex)
			}
		}
		// Empty Excel-authored MultiPages can retain a stale raw ListIndex. Keep
		// it on no-op replay; projection exposes the logical empty selection -1.
		if collectionCount > 0 && (strip.SelectedIndex < 0 || int(strip.SelectedIndex) >= collectionCount) {
			return nil, templateError(current, Invalid, desired.Name, "selectedIndex", "selection outside final collection")
		}
		edit := oforms.TabEdit{Control: desired.Name, State: strip}
		oldProps, err := propertyMap(old.Properties)
		if err != nil {
			return nil, err
		}
		desiredProps, err := propertyMap(desired.Properties)
		if err != nil {
			return nil, err
		}
		if owner.MultiPage != nil && !reflect.DeepEqual(oldProps["enabled"], desiredProps["enabled"]) {
			enabled, ok := desiredProps["enabled"].(bool)
			if !ok {
				return nil, templateError(current, Invalid, desired.Name, "enabled", "requires Boolean")
			}
			if !reflect.DeepEqual(old.Enabled, desired.Enabled) && desired.Enabled != nil && *desired.Enabled != enabled {
				return nil, templateError(current, Conflict, desired.Name, "enabled", "conflicting enabled aliases")
			}
			if oldProps["enabled"] != nil && baseline.Enabled != nil && oldProps["enabled"] != *baseline.Enabled {
				return nil, templateError(current, Stale, desired.Name, "enabled", "before enabled is stale")
			}
			edit.Enabled = new(enabled)
		}
		if owner.MultiPage != nil && !reflect.DeepEqual(old.Enabled, desired.Enabled) && desired.Enabled != nil {
			if retained && old.Enabled != nil && !reflect.DeepEqual(old.Enabled, baseline.Enabled) {
				return nil, templateError(current, Stale, desired.Name, "enabled", "before enabled is stale")
			}
			edit.Enabled = desired.Enabled
		}
		edits = append(edits, edit)
	}
	result, err := oforms.ApplyTabEdits(current, edits, codePage)
	if err != nil {
		return nil, templateError(current, Invalid, "", "tabs", fmt.Sprint(err))
	}
	return result, nil
}

// Template authoring carries explicit selection intent even when its numeric
// value matches the old index and topology has changed the selected identity.
func applyExplicitTabSelection(form *oforms.Form, desired spec.FormSpec, codePage uint16) (*oforms.Form, error) {
	var edits []oforms.TabEdit
	for _, control := range desired.Controls {
		if control.SelectedIndex == nil || (!strings.EqualFold(control.Type, "MultiPage") && !strings.EqualFold(control.Type, "TabStrip")) {
			continue
		}
		owner := findModelControl(form.Controls, control.Name)
		if owner == nil {
			return nil, templateError(form, Invalid, control.Name, "selectedIndex", "owner missing")
		}
		var state *oforms.TabStrip
		if owner.MultiPage != nil {
			state = owner.MultiPage.Hidden.TabStrip
		} else {
			state = owner.TabStrip
		}
		next := *state
		index := *control.SelectedIndex
		count := len(state.Tabs)
		if owner.MultiPage != nil {
			count = len(owner.MultiPage.Pages)
		}
		if index < -1 || index > math.MaxInt32 || count == 0 && index != -1 || count > 0 && (index < 0 || index >= count) {
			return nil, templateError(form, Invalid, control.Name, "selectedIndex", "selection outside final collection")
		}
		// Logical empty -1 does not rewrite a stale saved wire value on replay.
		if count > 0 {
			next.SelectedIndex = int32(index)
		}
		edits = append(edits, oforms.TabEdit{Control: control.Name, State: &next})
	}
	if len(edits) == 0 {
		return form, nil
	}
	result, err := oforms.ApplyTabEdits(form, edits, codePage)
	if err != nil {
		return nil, templateError(form, Invalid, "", "selectedIndex", err.Error())
	}
	return result, nil
}

// Keep a source tab's omitted attributes from being converted into reset intent.
func overlayTabs(previous, desired []spec.FormSpecTab) []spec.FormSpecTab {
	old := map[string]spec.FormSpecTab{}
	for _, tab := range previous {
		old[tab.Name] = tab
	}
	next := make([]spec.FormSpecTab, 0, len(desired))
	for _, input := range desired {
		tab, ok := old[input.Name]
		if !ok {
			tab = input
		} else {
			for _, field := range []struct {
				target **string
				source *string
			}{{&tab.Caption, input.Caption}, {&tab.ControlTipText, input.ControlTipText}, {&tab.Tag, input.Tag}, {&tab.Accelerator, input.Accelerator}} {
				if field.source != nil {
					*field.target = new(*field.source)
				}
			}
			if input.Enabled != nil {
				tab.Enabled = new(*input.Enabled)
			}
			if input.Visible != nil {
				tab.Visible = new(*input.Visible)
			}
		}
		next = append(next, tab)
	}
	return next
}
