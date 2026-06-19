package vbaproject

import (
	"fmt"
	"strings"
)

// NormalizeModuleSource converts a disk-form source (.bas/.cls, as exported to
// a file) into the in-bin form that Excel keeps in the module stream (Module.Source).
// It is a pure string-to-string transform, independent of encoding: disk form
// is UTF-8, and the conversion to the in-bin MBCS encoding is handled by
// encodeMBCS in Write.
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
		return "", fmt.Errorf("vbaproject: form modules are not supported in v1; handle externally")
	default:
		return "", fmt.Errorf("vbaproject: unsupported ModuleType %d", mt)
	}
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
