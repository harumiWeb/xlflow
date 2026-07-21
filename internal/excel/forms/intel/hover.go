package intel

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/excel/forms"
)

// Hover is protocol-neutral documentation for a UserForm YAML token.
type Hover struct {
	Contents string
	Range    Range
}

// HoverYAML returns contract-backed documentation for a YAML field key or a
// control type / ProgID scalar. It deliberately returns nil for unknown YAML
// so callers can preserve the normal LSP no-hover behavior.
func HoverYAML(source string, pos Position) *Hover {
	doc := ParseYAML(source)
	if doc == nil {
		return nil
	}
	for path, field := range doc.Fields {
		if contains(field.KeyRange, pos) {
			if property, scope, ok := propertyAtPath(doc, path); ok {
				return &Hover{Contents: renderPropertyHover(lastPathSegment(path), property, scope), Range: field.KeyRange}
			}
		}
		if contains(field.ValueRange, pos) {
			switch strings.ToLower(lastPathSegment(path)) {
			case "type":
				return controlTypeHover(doc, parentPath(path), field)
			case "progid":
				return progIDHover(field)
			}
		}
	}
	return nil
}

func propertyAtPath(doc *Document, path string) (forms.PropertyContract, string, bool) {
	contract := forms.UserFormContract()
	parent := parentPath(path)
	name := lastPathSegment(path)
	switch {
	case parent == "":
		property, ok := contract.DocumentProperties[name]
		return property, "document root", ok
	case parent == "form":
		property, ok := contract.FormProperties[name]
		return property, "form", ok
	case isControlPath(parent):
		control := controlAtPath(doc.Source, parent)
		if property, ok := forms.LookupControlProperty(control.Type, name); ok {
			if _, common := contract.CommonControlProperties[name]; common {
				return property, "all controls", true
			}
			return property, control.Type, true
		}
	}
	return forms.PropertyContract{}, "", false
}

func controlTypeHover(doc *Document, controlPath string, field FieldNodes) *Hover {
	typeName := strings.TrimSpace(field.Value.Value)
	if control, ok := forms.LookupControlContract(typeName); ok {
		contents := []string{
			fmt.Sprintf("### `%s`", control.Type),
			"**Kind:** built-in UserForm control  ",
			"**ProgID:** `" + control.ProgID + "`  ",
			"**Container:** " + yesNo(control.CanContainChildren),
		}
		if control.CanContainChildren {
			contents = append(contents, "May be used as a `parentId` target.")
		} else {
			contents = append(contents, "Cannot be used as a `parentId` target.")
		}
		return &Hover{Contents: strings.Join(contents, "\n"), Range: field.ValueRange}
	}
	if control := controlAtPath(doc.Source, controlPath); strings.TrimSpace(control.progID) != "" {
		return &Hover{Contents: customControlHover(control.progID), Range: field.ValueRange}
	}
	return nil
}

func progIDHover(field FieldNodes) *Hover {
	progID := strings.TrimSpace(field.Value.Value)
	if control, ok := forms.LookupControlContractByProgID(progID); ok {
		contents := []string{
			fmt.Sprintf("### `%s`", control.ProgID),
			"**Built-in control:** `" + control.Type + "`  ",
			"**Support:** supported",
			"Creates the built-in `" + control.Type + "` control when used with the matching `type`.",
		}
		return &Hover{Contents: strings.Join(contents, "\n"), Range: field.ValueRange}
	}
	if progID != "" {
		return &Hover{Contents: customControlHover(progID), Range: field.ValueRange}
	}
	return nil
}

func renderPropertyHover(name string, property forms.PropertyContract, scope string) string {
	lines := []string{fmt.Sprintf("### `%s`", name)}
	lines = append(lines, "**Type:** "+string(property.ValueType)+"  ")
	if property.Required {
		lines = append(lines, "**Required:** yes  ")
	} else {
		lines = append(lines, "**Required:** no  ")
	}
	if len(property.ApplicableControls) > 0 {
		lines = append(lines, "**Applies to:** `"+strings.Join(property.ApplicableControls, "`, `")+"`  ")
	} else {
		lines = append(lines, "**Applies to:** "+scope+"  ")
	}
	lines = append(lines, "**Support:** "+string(property.SupportLevel), "", property.Description)
	if limitation := supportLimitation(name, property.SupportLevel); limitation != "" {
		lines = append(lines, "", limitation)
	}
	return strings.Join(lines, "\n")
}

func supportLimitation(name string, level forms.SupportLevel) string {
	switch level {
	case forms.SupportLevelBestEffort:
		return "xlflow attempts to apply this during build; inspect the rebuilt Designer to confirm the result."
	case forms.SupportLevelObservedOnly:
		return "This is captured design-time state. xlflow may apply it during build on a best-effort basis, but Excel Designer round-trips are not guaranteed."
	case forms.SupportLevelSnapshotOnly:
		return "Snapshot-oriented metadata captured from Excel. It is not a guaranteed normal build input for hand-authored specifications."
	case forms.SupportLevelCustomUnchecked:
		if name == "properties" {
			return "This is not an unrestricted build escape hatch: xlflow does not validate arbitrary custom properties or guarantee that Designer creation applies them."
		}
		return "Accepted for compatibility without type-specific validation or a guaranteed build result."
	default:
		return ""
	}
}

func customControlHover(progID string) string {
	return strings.Join([]string{
		"### `" + progID + "`",
		"**Kind:** custom ProgID  ",
		"**Support:** custom/unchecked",
		"xlflow may attempt to create this control and validates common structural fields. Type-specific property validation is unavailable, and Designer compatibility depends on the installed control.",
	}, "\n")
}

func yesNo(value bool) string {
	if value {
		return "yes"
	}
	return "no"
}
