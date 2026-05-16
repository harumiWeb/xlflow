package forms

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/lint"
)

type UserFormArtifactIssue struct {
	FormName   string
	Path       string
	Line       int
	Message    string
	Suggestion string
}

func (i UserFormArtifactIssue) LintIssue(formsDir string) lint.Issue {
	relPath := i.Path
	if strings.TrimSpace(formsDir) != "" {
		if rel, err := filepath.Rel(formsDir, i.Path); err == nil {
			relPath = filepath.ToSlash(filepath.Join("src/forms", rel))
		} else {
			relPath = filepath.ToSlash(i.Path)
		}
	}
	return lint.Issue{
		Code:       "FRM201",
		Severity:   "warning",
		File:       relPath,
		Line:       i.Line,
		Message:    i.Message,
		Kind:       "user_form",
		Symbol:     "form artifact",
		Suggestion: i.Suggestion,
	}
}

func ValidateUserFormArtifactsAgainstSpecs(formsDir string, targetForms map[string]bool) ([]UserFormArtifactIssue, error) {
	if strings.TrimSpace(formsDir) == "" {
		return nil, nil
	}
	if _, err := os.Stat(formsDir); err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}

	specDir := filepath.Join(formsDir, "specs")
	if _, err := os.Stat(specDir); err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}

	frmIndex, err := collectUserFormArtifacts(formsDir)
	if err != nil {
		return nil, err
	}

	issues := make([]UserFormArtifactIssue, 0)
	err = filepath.WalkDir(specDir, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		format, formatErr := snapshotFormatFromPath(path)
		if formatErr != nil {
			return nil
		}
		displayPath, displayErr := filepath.Rel(formsDir, path)
		if displayErr != nil {
			displayPath = path
		}
		spec, specErr := LoadFormSpec(SpecInput{Path: path, DisplayPath: filepath.ToSlash(displayPath), Format: format})
		if specErr != nil {
			return specErr
		}

		specFileBase := strings.TrimSpace(strings.TrimSuffix(d.Name(), filepath.Ext(d.Name())))
		formName := strings.TrimSpace(spec.Form.Name)
		if len(targetForms) > 0 && !targetForms[formName] && !targetForms[specFileBase] {
			return nil
		}

		if specFileBase != "" && !strings.EqualFold(specFileBase, formName) {
			issues = append(issues, UserFormArtifactIssue{
				FormName:   formName,
				Path:       path,
				Message:    fmt.Sprintf("UserForm spec file %q declares form.name %q. Push resolves Designer-backed forms from .frm artifact names, so the spec filename and form.name must match.", d.Name(), formName),
				Suggestion: "Rename the spec file or update form.name so both use the same UserForm name, then rebuild before push.",
			})
		}

		matches := frmIndex[strings.ToLower(formName)]
		if len(matches) == 0 {
			issues = append(issues, UserFormArtifactIssue{
				FormName:   formName,
				Path:       path,
				Message:    fmt.Sprintf("UserForm spec %q has no matching .frm artifact for form.name %q under src/forms. Push would ignore the spec and operate on stale workbook source.", d.Name(), formName),
				Suggestion: "Run xlflow form build on this spec first so src/forms/<FormName>.frm exists and matches the spec before push.",
			})
			return nil
		}
		if len(matches) > 1 {
			paths := make([]string, 0, len(matches))
			for _, match := range matches {
				if rel, relErr := filepath.Rel(formsDir, match); relErr == nil {
					paths = append(paths, filepath.ToSlash(filepath.Join("src/forms", rel)))
				} else {
					paths = append(paths, filepath.ToSlash(match))
				}
			}
			sort.Strings(paths)
			issues = append(issues, UserFormArtifactIssue{
				FormName:   formName,
				Path:       path,
				Message:    fmt.Sprintf("UserForm spec %q matched multiple .frm artifacts for form.name %q: %s.", d.Name(), formName, strings.Join(paths, ", ")),
				Suggestion: "Keep exactly one .frm artifact per spec form.name, then rebuild before push.",
			})
			return nil
		}

		frmPath := matches[0]
		vbName, line, nameErr := readUserFormArtifactVBName(frmPath)
		if nameErr != nil {
			return nameErr
		}
		if !strings.EqualFold(vbName, formName) {
			issues = append(issues, UserFormArtifactIssue{
				FormName:   formName,
				Path:       frmPath,
				Line:       line,
				Message:    fmt.Sprintf("UserForm artifact %q declares Attribute VB_Name = %q, but the matching spec requires form.name %q. Push would import the artifact under the wrong Designer name.", filepath.Base(frmPath), vbName, formName),
				Suggestion: "Rebuild the form from the spec so the exported .frm header and filename both match form.name before push.",
			})
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return issues, nil
}

func collectUserFormArtifacts(formsDir string) (map[string][]string, error) {
	index := make(map[string][]string)
	specDir := filepath.Join(formsDir, "specs")
	codeDir := filepath.Join(formsDir, "code")
	err := filepath.WalkDir(formsDir, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			if samePath(path, specDir) || samePath(path, codeDir) {
				return filepath.SkipDir
			}
			return nil
		}
		if !strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			return nil
		}
		name := strings.ToLower(strings.TrimSpace(strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))))
		if name == "" {
			return nil
		}
		index[name] = append(index[name], path)
		return nil
	})
	if err != nil {
		return nil, err
	}
	for name := range index {
		sort.Strings(index[name])
	}
	return index, nil
}

func readUserFormArtifactVBName(path string) (string, int, error) {
	body, err := os.ReadFile(path)
	if err != nil {
		return "", 0, fmt.Errorf("read %s: %w", path, err)
	}
	lines := splitNormalizedLines(string(body))
	for idx, line := range lines {
		trimmed := strings.TrimSpace(line)
		if !strings.HasPrefix(strings.ToLower(trimmed), "attribute vb_name") {
			continue
		}
		parts := strings.SplitN(trimmed, "=", 2)
		if len(parts) != 2 {
			break
		}
		name := strings.TrimSpace(parts[1])
		name = strings.Trim(name, "\"")
		if name == "" {
			break
		}
		return name, idx + 1, nil
	}
	return "", 0, fmt.Errorf("read %s: Attribute VB_Name was not found", path)
}
