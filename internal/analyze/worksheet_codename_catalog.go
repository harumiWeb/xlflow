package analyze

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/ooxml"
)

func loadWorksheetCodeNameCatalog(rootDir, workbookPath string) (WorksheetCodeNameCatalog, map[string]any) {
	workbookPath = strings.TrimSpace(workbookPath)
	if workbookPath == "" {
		return nil, worksheetCodeNameCapabilityWarning("", "the configured workbook path is empty")
	}
	if !filepath.IsAbs(workbookPath) {
		workbookPath = filepath.Join(rootDir, workbookPath)
	}
	pkg, err := ooxml.Open(workbookPath)
	if err != nil {
		return nil, worksheetCodeNameCapabilityWarning(workbookPath, err.Error())
	}
	defer func() { _ = pkg.Close() }()
	catalog, err := pkg.ReadWorksheetCodeNameCatalog()
	if err != nil {
		return nil, worksheetCodeNameCapabilityWarning(workbookPath, err.Error())
	}
	if issues := catalog.Issues(); len(issues) > 0 {
		return catalog, worksheetCodeNameCapabilityWarning(workbookPath, fmt.Sprintf("the workbook catalog excluded %d malformed or ambiguous worksheet entries", len(issues)))
	}
	return catalog, nil
}

func worksheetCodeNameCapabilityWarning(workbookPath, detail string) map[string]any {
	message := "Worksheet CodeName metadata is unavailable; VBA260 was skipped."
	if detail != "" {
		message = fmt.Sprintf("%s %s", message, detail)
	}
	return map[string]any{
		"code":       "analysis_capability_unavailable",
		"capability": "worksheet_codename_catalog",
		"file":       workbookPath,
		"rules":      []string{"VBA260"},
		"message":    message,
	}
}
