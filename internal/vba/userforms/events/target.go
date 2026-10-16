package events

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// Target identifies the VBA object whose events are exposed for a form or
// control. A Page keeps its target ID but uses its parent MultiPage object.
type Target struct {
	ID         string
	ObjectName string
	Type       string
}

// ResolveTarget resolves the root form (empty id) or a built-in control. Page
// targets resolve to their parent MultiPage because the parent owns the Page
// event connection point in VBA.
func ResolveTarget(form spec.FormSpec, targetID string) (Target, error) {
	if targetID == "" {
		if strings.TrimSpace(form.Form.Name) == "" {
			return Target{}, fmt.Errorf("UserForm name is missing")
		}
		// The UserForm code module uses the VBE's fixed UserForm event object
		// prefix regardless of the component's form.name.
		return Target{ObjectName: "UserForm", Type: "UserForm"}, nil
	}

	controls := make(map[string]spec.FormSpecControl, len(form.Controls))
	for _, control := range form.Controls {
		controls[control.ID] = control
	}
	control, ok := controls[targetID]
	if !ok {
		return Target{}, fmt.Errorf("unknown UserForm control target %q", targetID)
	}
	if control.ProgID != "" && !spec.ProgIDMatchesControlType(control.Type, control.ProgID) {
		return Target{}, fmt.Errorf("control %q has a ProgID that does not match its built-in type", control.Name)
	}

	typeName := strings.TrimSpace(control.Type)
	if strings.EqualFold(typeName, "Page") {
		parent, found := controls[control.ParentID]
		if !found || !strings.EqualFold(strings.TrimSpace(parent.Type), "MultiPage") {
			return Target{}, fmt.Errorf("page target %q must have a built-in MultiPage parent", control.Name)
		}
		if parent.ProgID != "" && !spec.ProgIDMatchesControlType(parent.Type, parent.ProgID) {
			return Target{}, fmt.Errorf("page parent %q has a ProgID that does not match MultiPage", parent.Name)
		}
		if _, supported := ForType(parent.Type, parent.Name); !supported {
			return Target{}, fmt.Errorf("page parent %q has unsupported type %q", parent.Name, parent.Type)
		}
		return Target{ID: targetID, ObjectName: parent.Name, Type: parent.Type}, nil
	}
	if _, supported := ForType(typeName, control.Name); !supported {
		return Target{}, fmt.Errorf("unsupported UserForm control type %q", control.Type)
	}
	if strings.TrimSpace(control.Name) == "" {
		return Target{}, fmt.Errorf("UserForm control name is missing")
	}
	return Target{ID: targetID, ObjectName: control.Name, Type: typeName}, nil
}
