package vbaproject

import (
	"fmt"
	"strings"
)

// NormalizeModuleSource converts a disk-form source (.bas/.cls, as exported to
// a file) into the in-bin form that Excel keeps in the module stream (Module.Source).
// It is a pure string-to-string transform, independent of encoding: disk form
// is UTF-8, and the conversion to the in-bin MBCS encoding is handled by
// ovba.EncodeMBCS in Write.
//
// existing is required only for document modules (it supplies the Attribute
// header of the existing in-bin module). It may be nil for std/class modules.
func NormalizeModuleSource(mt ModuleType, disk string, existing *Module) (string, error) {
	disk = toCRLF(disk)
	switch mt {
	case ModuleStd:
		return normalizeStd(disk)
	case ModuleClass:
		return normalizeClass(disk)
	case ModuleDocument:
		return normalizeDocument(disk, existing)
	case ModuleForm:
		return normalizeForm(disk, existing)
	default:
		return "", fmt.Errorf("vbaproject: unsupported ModuleType %d", mt)
	}
}

// ExportModuleSource converts the source stored in a vbaProject.bin module
// stream into xlflow's tracked, UTF-8 disk representation. UserForms require
// their separate Designer export and are intentionally rejected here.
func ExportModuleSource(module Module) (string, error) {
	source := toLF(module.Source)
	switch module.Type {
	case ModuleStd:
		if err := ValidateModuleIdentity(module.Name, source); err != nil {
			return "", err
		}
		return ensureFinalNewline(source), nil
	case ModuleClass:
		if err := ValidateModuleIdentity(module.Name, source); err != nil {
			return "", err
		}
		body := removeInternalClassAttributes(source)
		header := "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
		return header + ensureFinalNewline(body), nil
	case ModuleDocument:
		return exportDocumentSource(source)
	case ModuleForm:
		return "", fmt.Errorf("vbaproject: UserForm %q requires a Designer export", module.Name)
	default:
		return "", fmt.Errorf("vbaproject: unsupported ModuleType %d", module.Type)
	}
}

func exportDocumentSource(source string) (string, error) {
	lines := strings.Split(source, "\n")
	firstBody := 0
	for firstBody < len(lines) && strings.HasPrefix(lines[firstBody], "Attribute ") {
		firstBody++
	}
	if firstBody == 0 {
		return "", fmt.Errorf("vbaproject: document module has no Attribute header")
	}
	body := strings.Join(lines[firstBody:], "\n")
	body = strings.TrimPrefix(body, "\n")
	return ensureFinalNewline(body), nil
}

func removeInternalClassAttributes(source string) string {
	lines := strings.Split(source, "\n")
	out := lines[:0]
	for _, line := range lines {
		switch {
		case strings.HasPrefix(line, "Attribute VB_Base "):
		case strings.HasPrefix(line, "Attribute VB_TemplateDerived "):
		case strings.HasPrefix(line, "Attribute VB_Customizable "):
		default:
			out = append(out, line)
		}
	}
	return strings.Join(out, "\n")
}

func toLF(source string) string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	return strings.ReplaceAll(source, "\r", "\n")
}

func ensureFinalNewline(source string) string {
	if strings.HasSuffix(source, "\n") {
		return source
	}
	return source + "\n"
}

// ValidateModuleIdentity verifies that the in-bin Attribute VB_Name agrees
// with the component identity used by PROJECT, dir, and the VBA stream. It is
// used whenever pack updates or creates a standard or class component. The
// textual source remains authoritative for code, but its declared identity
// must agree with the component model.
func ValidateModuleIdentity(expectedName, source string) error {
	for line := range strings.SplitSeq(toCRLF(source), "\r\n") {
		if !strings.HasPrefix(line, "Attribute VB_Name ") {
			continue
		}
		_, value, ok := strings.Cut(line, "=")
		if !ok {
			break
		}
		value = strings.TrimSpace(value)
		if len(value) < 2 || value[0] != '"' || value[len(value)-1] != '"' {
			return fmt.Errorf("vbaproject: Attribute VB_Name is malformed")
		}
		declared := value[1 : len(value)-1]
		if declared != expectedName {
			return fmt.Errorf("vbaproject: Attribute VB_Name %q does not match component name %q", declared, expectedName)
		}
		return nil
	}
	return fmt.Errorf("vbaproject: Attribute VB_Name was not found")
}

// toCRLF normalizes mixed line endings to CRLF (the in-bin form is always CRLF). Observed disk
// forms are already CRLF, so it is effectively a no-op, but it guards against stray LFs.
func toCRLF(s string) string {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	return strings.ReplaceAll(s, "\n", "\r\n")
}

// hasAttrLine reports whether any line of s starts with `Attribute <name> ` (a line anchor).
// strings.Contains could match a string literal in the code body, so this matches at the line start.
func hasAttrLine(s, name string) bool {
	prefix := "Attribute " + name + " "
	for _, ln := range strings.Split(s, "\r\n") {
		if strings.HasPrefix(ln, prefix) {
			return true
		}
	}
	return false
}

// classVBBaseGUID is the default VB_Base value of an ordinary class (generic and host-independent;
// measured to differ from the document ThisWorkbook/Sheet host GUIDs). Including the leading "0", it
// matches the in-bin `Attribute VB_Base = "0{...}"` string byte-for-byte (reproducing exactly what VBE emits).
const classVBBaseGUID = `0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}`

const (
	workbookVBBaseGUID  = `0{00020819-0000-0000-C000-000000000046}`
	worksheetVBBaseGUID = `0{00020820-0000-0000-C000-000000000046}`
)

// NewDocumentModule returns a canonical, source-only Excel document module.
// worksheet selects the Excel Worksheet host class; false selects Workbook.
// The disk source must contain code only, matching the source-tree contract.
func NewDocumentModule(name, disk string, worksheet bool) (Module, error) {
	base := workbookVBBaseGUID
	if worksheet {
		base = worksheetVBBaseGUID
	}
	header := strings.Join([]string{
		`Attribute VB_Name = "` + name + `"`,
		`Attribute VB_Base = "` + base + `"`,
		"Attribute VB_GlobalNameSpace = False",
		"Attribute VB_Creatable = False",
		"Attribute VB_PredeclaredId = True",
		"Attribute VB_Exposed = True",
		"Attribute VB_TemplateDerived = False",
		"Attribute VB_Customizable = True",
	}, "\r\n") + "\r\n"
	existing := Module{Name: name, StreamName: name, Type: ModuleDocument, Source: header}
	source, err := NormalizeModuleSource(ModuleDocument, disk, &existing)
	if err != nil {
		return Module{}, err
	}
	existing.Source = source
	return existing, nil
}

// normalizeClass removes the leading VERSION..END block of the disk .cls and injects the
// VB_Base / VB_TemplateDerived / VB_Customizable that the in-bin form carries, at their proper positions.
func normalizeClass(disk string) (string, error) {
	lines := strings.Split(disk, "\r\n")
	if !strings.HasPrefix(lines[0], "VERSION ") {
		return "", fmt.Errorf("vbaproject: class disk form does not start with a VERSION header")
	}
	endIdx := -1
	for i, ln := range lines {
		if ln == "END" {
			endIdx = i
			break
		}
	}
	if endIdx < 0 {
		return "", fmt.Errorf("vbaproject: class disk form has no END line")
	}
	body := lines[endIdx+1:] // drop VERSION..END, keep everything from Attribute onward (preserving the trailing empty element)

	var sawName, sawExposed bool
	out := make([]string, 0, len(body)+3)
	for _, ln := range body {
		out = append(out, ln)
		if strings.HasPrefix(ln, "Attribute VB_Name ") {
			out = append(out, `Attribute VB_Base = "`+classVBBaseGUID+`"`)
			sawName = true
		}
		if strings.HasPrefix(ln, "Attribute VB_Exposed ") {
			out = append(out, "Attribute VB_TemplateDerived = False")
			out = append(out, "Attribute VB_Customizable = False")
			sawExposed = true
		}
	}
	if !sawName || !sawExposed {
		return "", fmt.Errorf("vbaproject: class disk form is missing the injection anchors (VB_Name/VB_Exposed)")
	}
	return strings.Join(out, "\r\n"), nil
}

// normalizeDocument: the disk form is pure code with zero attributes (ThisWorkbook/Sheet). Since the
// attributes cannot be reconstructed from disk, it preserves the leading run of Attribute lines from
// existing (the in-bin form read in) and prepends them to the disk code. Host GUIDs
// (ThisWorkbook=00020819 / Sheet=00020820), VB_PredeclaredId, etc. come from existing rather than being
// hardcoded (the policy for documents is to replace only the code body).
func normalizeDocument(disk string, existing *Module) (string, error) {
	if existing == nil {
		return "", fmt.Errorf("vbaproject: document module requires existing (the in-bin header)")
	}
	existingLines := strings.Split(toCRLF(existing.Source), "\r\n")
	var header []string
	for _, ln := range existingLines {
		if strings.HasPrefix(ln, "Attribute ") {
			header = append(header, ln)
		} else {
			break // once the leading Attribute block ends, the code body begins
		}
	}
	if len(header) == 0 {
		return "", fmt.Errorf("vbaproject: existing.Source has no Attribute header")
	}
	// disk is assumed to be already toCRLF'd by NormalizeModuleSource and to have a trailing \r\n
	// (a different origin than normalizeClass, which gets its trailing CRLF from Split's empty element).
	return strings.Join(header, "\r\n") + "\r\n" + disk, nil
}

// normalizeStd: the disk .bas is almost identical to the in-bin form (Attribute VB_Name + code), so this is an identity transform.
// A missing VB_Name is treated as a sign of a broken disk form and returns an error (no silent progress).
func normalizeStd(disk string) (string, error) {
	if !hasAttrLine(disk, "VB_Name") {
		return "", fmt.Errorf("vbaproject: std disk form has no Attribute VB_Name")
	}
	return disk, nil
}

// normalizeForm produces the in-bin source for a UserForm code-behind update. A disk .frm carries a
// VERSION / Begin..End designer block, then Attribute VB_* lines, then code. Only the code-behind is
// editable here: the in-bin Attribute header (which includes the form-specific VB_Base linking the
// module to its designer storage) is taken from existing, exactly as normalizeDocument does, so the
// rewritten code module stays bound to the template's carried designer storage even if the disk .frm's
// own attributes differ. The code-behind boundary — everything after the last `Attribute VB_` line —
// mirrors internal/excel/forms.splitUserFormFRMSections, the canonical .frm splitter used by push/pull;
// it is reimplemented here to keep internal/pack self-contained (ADR-0012) and the two must stay in agreement.
func normalizeForm(disk string, existing *Module) (string, error) {
	if existing == nil {
		return "", fmt.Errorf("vbaproject: form module requires existing (the in-bin header)")
	}
	var header []string
	for _, ln := range strings.Split(toCRLF(existing.Source), "\r\n") {
		if strings.HasPrefix(ln, "Attribute ") {
			header = append(header, ln)
		} else {
			break // the leading Attribute block ends; the code body begins
		}
	}
	if len(header) == 0 {
		return "", fmt.Errorf("vbaproject: existing form module has no Attribute header")
	}
	diskLines := strings.Split(disk, "\r\n")
	lastAttr := -1
	for i, ln := range diskLines {
		if strings.HasPrefix(strings.TrimSpace(ln), "Attribute VB_") {
			lastAttr = i
		}
	}
	if lastAttr < 0 {
		return "", fmt.Errorf("vbaproject: form disk form (.frm) has no Attribute VB_ line")
	}
	start := lastAttr + 1
	for start < len(diskLines) && strings.TrimSpace(diskLines[start]) == "" {
		start++ // skip blank lines between the attribute block and the code body
	}
	out := append([]string{}, header...)
	out = append(out, diskLines[start:]...)
	return strings.Join(out, "\r\n"), nil
}
