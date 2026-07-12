package workbookformat

import (
	"fmt"
	"path/filepath"
	"strings"
)

const (
	ExtXLSM = ".xlsm"
	ExtXLAM = ".xlam"
	ExtXLSB = ".xlsb"
	ExtXLSX = ".xlsx"
	ExtXLTX = ".xltx"
	ExtXLTM = ".xltm"
)

const UnsupportedErrorCode = "workbook_format_unsupported"

type UnsupportedError struct {
	Capability string
	Extension  string
}

func (e UnsupportedError) Error() string {
	ext := strings.ToLower(strings.TrimSpace(e.Extension))
	if ext == "" {
		ext = "<none>"
	}
	capability := strings.TrimSpace(e.Capability)
	if capability == "" {
		capability = "workbook operation"
	}
	return fmt.Sprintf("xlflow %s does not currently support %s workbooks", capability, ext)
}

func Format(path string) string {
	ext := strings.ToLower(filepath.Ext(path))
	return strings.TrimPrefix(ext, ".")
}

func NormalizeProjectWorkbookName(name string) (string, error) {
	name = strings.TrimSpace(name)
	if name == "" {
		return "Book.xlsm", nil
	}
	name = filepath.Base(name)
	ext := strings.ToLower(filepath.Ext(name))
	if ext == "" {
		return name + ExtXLSM, nil
	}
	if !IsProjectWorkbookExt(ext) {
		return "", fmt.Errorf("workbook name must use .xlsm, .xlam, or .xlsb extension: %s", name)
	}
	return name, nil
}

func ValidateProjectWorkbookPath(path string) error {
	ext := strings.ToLower(filepath.Ext(path))
	if !IsProjectWorkbookExt(ext) {
		return fmt.Errorf("workbook path must use .xlsm, .xlam, or .xlsb extension: %s", path)
	}
	return nil
}

func IsProjectWorkbookExt(ext string) bool {
	switch strings.ToLower(ext) {
	case ExtXLSM, ExtXLAM, ExtXLSB:
		return true
	default:
		return false
	}
}

func ValidateFormulaSnapshotWorkbook(path string) error {
	switch strings.ToLower(filepath.Ext(path)) {
	case ExtXLSX, ExtXLSM:
		return nil
	case ExtXLSB:
		return UnsupportedError{Capability: "formulas pull", Extension: ExtXLSB}
	default:
		return fmt.Errorf("source workbook must end in .xlsx or .xlsm: %s", path)
	}
}

func ValidateFileInspectWorkbook(path, capability string) error {
	switch strings.ToLower(filepath.Ext(path)) {
	case ExtXLSX, ExtXLSM, ExtXLTX, ExtXLTM:
		return nil
	case ExtXLSB:
		return UnsupportedError{Capability: capability, Extension: ExtXLSB}
	default:
		return fmt.Errorf("unsupported extension %q; expected .xlsx, .xlsm, .xltx, or .xltm", filepath.Ext(path))
	}
}

func ValidateDiffWorkbook(path string) error {
	return ValidateFileInspectWorkbook(path, "diff")
}

func ValidatePackTemplate(path string) error {
	if strings.EqualFold(filepath.Ext(path), ExtXLSB) {
		return UnsupportedError{Capability: "pack", Extension: ExtXLSB}
	}
	return nil
}
