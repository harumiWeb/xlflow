package filepush

import (
	"errors"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
)

// collectSources resolves the managed source tree into the pack source plan,
// applying the same source transforms the Excel bridge applies while staging
// import copies: UserForm sidecar merge, document-module normalization,
// folder annotation updates, and Erl line-number instrumentation.
func collectSources(root string, cfg config.Config) ([]packpkg.SourceModule, error) {
	components, err := sourceinventory.Discover(sourceinventory.Options{
		Root: root, Config: cfg, ValidateFormArtifacts: true,
	})
	if err != nil {
		var layoutErr *sourceinventory.LayoutError
		if errors.As(err, &layoutErr) {
			return nil, fmt.Errorf("%w: %v", packpkg.ErrAmbiguousLayout, err)
		}
		return nil, err
	}
	roots := resolvedRoots(root, cfg)
	sidecar := isSidecarMode(cfg.UserForm.CodeSource)
	lineNumbers := cfg.VBA.LineNumbers.Enabled
	sources := make([]packpkg.SourceModule, 0, len(components))
	for _, component := range components {
		if component.Type == sourceinventory.ComponentDocument && !strings.EqualFold(filepath.Ext(component.AbsolutePath), ".bas") {
			// The Excel bridge only applies document modules discovered as
			// <name>.bas under the workbook root; a .cls there is never
			// fingerprinted or imported. Skip it for the same outcome.
			continue
		}
		source := string(component.Source)
		sidecarText, hasSidecar := "", false
		if component.Type == sourceinventory.ComponentForm && sidecar {
			for _, artifact := range component.Related {
				if strings.EqualFold(filepath.Ext(artifact.Path), ".bas") && strings.EqualFold(filepath.Base(artifact.Path), component.Name+".bas") {
					sidecarText, hasSidecar = string(artifact.Source), true
					break
				}
			}
		}
		transformed, err := transformComponentSource(component, source, sidecarText, hasSidecar, roots, cfg, lineNumbers)
		if err != nil {
			return nil, err
		}
		sources = append(sources, packpkg.SourceModule{
			SourcePath:   component.SourcePath,
			RelatedPaths: component.RelatedPaths(),
			Name:         component.Name,
			Type:         packpkg.ModuleType(component.Type),
			Source:       transformed,
		})
	}
	return sources, nil
}

// transformComponentSource applies push-time text transforms in the same order
// the bridge applies them: for document modules the exported-attribute
// normalization runs first (UpdateDocumentModules), then the folder annotation
// update, then Erl instrumentation (PrepareSourceForImport). In sidecar mode a
// UserForm's code-behind comes from its code/<Name>.bas file verbatim (the
// bridge replaces the imported code module with the instrumented sidecar
// text), so the sidecar is numbered but never re-annotated.
func transformComponentSource(component sourceinventory.Component, source, sidecarText string, hasSidecar bool, roots sourceRoots, cfg config.Config, lineNumbers bool) (string, error) {
	if component.Type == sourceinventory.ComponentDocument {
		source = normalizeDocumentModuleContent(source)
	}
	rootDir := componentRootDir(component.Type, roots)
	if rootDir != "" {
		source = updateFolderAnnotationText(source, cfg.VBA.FolderAnnotation, folderAnnotationForPath(rootDir, component.AbsolutePath))
	}
	if lineNumbers {
		numbered, issue := tryAddLineNumbers(source)
		if issue != nil {
			return "", fmt.Errorf("%w: %s:%d: %s", ErrLineNumberSafety, component.SourcePath, issue.Line, issue.Message)
		}
		source = numbered
	}
	if !hasSidecar {
		return source, nil
	}
	code := sidecarText
	if lineNumbers {
		numbered, issue := tryAddLineNumbers(code)
		if issue != nil {
			return "", fmt.Errorf("%w: %s sidecar: %d: %s", ErrLineNumberSafety, component.SourcePath, issue.Line, issue.Message)
		}
		code = numbered
	}
	return forms.MergeUserFormCodeIntoFRM(source, code), nil
}

func componentRootDir(typ sourceinventory.ComponentType, roots sourceRoots) string {
	switch typ {
	case sourceinventory.ComponentStandard:
		return roots.modules
	case sourceinventory.ComponentClass:
		return roots.classes
	case sourceinventory.ComponentForm:
		return roots.forms
	case sourceinventory.ComponentDocument:
		return roots.workbook
	default:
		return ""
	}
}

// normalizeDocumentModuleContent ports VbaSourceHelper.NormalizeDocumentModuleContent:
// document-module disk sources are pure code, but hand-authored files may carry
// a class header or Attribute VB_* lines; both are dropped, and an empty body
// becomes the canonical "Option Explicit" module the bridge would produce.
func normalizeDocumentModuleContent(text string) string {
	lines := splitLinesKeepEmpty(text)
	filtered := make([]string, 0, len(lines))
	inClassHeader := false
	var classHeaderBuffer []string
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "VERSION 1.0 CLASS" {
			inClassHeader = true
			classHeaderBuffer = []string{line}
			continue
		}
		if inClassHeader {
			classHeaderBuffer = append(classHeaderBuffer, line)
			if trimmed == "END" {
				inClassHeader = false
				classHeaderBuffer = nil
			}
			continue
		}
		if attributeVBPattern.MatchString(trimmed) {
			continue
		}
		filtered = append(filtered, line)
	}
	if inClassHeader && len(classHeaderBuffer) > 0 {
		filtered = append(filtered, classHeaderBuffer...)
	}

	hasOptionExplicit := false
	hasNonHeaderCode := false
	for _, line := range filtered {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		if strings.EqualFold(trimmed, "Option Explicit") {
			hasOptionExplicit = true
			continue
		}
		hasNonHeaderCode = true
	}
	if !hasOptionExplicit && !hasNonHeaderCode {
		filtered = []string{"Option Explicit"}
	}
	return strings.Join(filtered, "\r\n")
}
