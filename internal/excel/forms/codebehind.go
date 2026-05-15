package forms

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

type UserFormCodeConflict struct {
	FormName    string
	FormPath    string
	SidecarPath string
}

func SyncUserFormCodeSidecars(formsDir string, targetForms map[string]bool) ([]UserFormCodeConflict, error) {
	if strings.TrimSpace(formsDir) == "" {
		return nil, nil
	}
	if _, err := os.Stat(formsDir); err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}
	sidecarDir := filepath.Join(formsDir, "code")
	updated := make([]UserFormCodeConflict, 0)
	err := filepath.WalkDir(formsDir, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			if samePath(path, sidecarDir) {
				return filepath.SkipDir
			}
			return nil
		}
		if !strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			return nil
		}
		formName := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if len(targetForms) > 0 && !targetForms[formName] {
			return nil
		}
		sidecarPath := filepath.Join(sidecarDir, formName+".bas")
		if _, err := os.Stat(sidecarPath); err != nil {
			if os.IsNotExist(err) {
				return nil
			}
			return err
		}
		formBody, err := os.ReadFile(path)
		if err != nil {
			return fmt.Errorf("read %s: %w", path, err)
		}
		sidecarBody, err := os.ReadFile(sidecarPath)
		if err != nil {
			return fmt.Errorf("read %s: %w", sidecarPath, err)
		}
		frmCode := NormalizeUserFormCodeText(ExtractUserFormCodeFromFRM(string(formBody)))
		sidecarCode := NormalizeUserFormCodeText(string(sidecarBody))
		if frmCode == sidecarCode {
			return nil
		}
		merged := MergeUserFormCodeIntoFRM(string(formBody), string(sidecarBody))
		if err := os.WriteFile(path, []byte(merged), 0o644); err != nil {
			return fmt.Errorf("write %s: %w", path, err)
		}
		updated = append(updated, UserFormCodeConflict{
			FormName:    formName,
			FormPath:    path,
			SidecarPath: sidecarPath,
		})
		return nil
	})
	if err != nil {
		return nil, err
	}
	return updated, nil
}

func FindUserFormCodeConflicts(formsDir string, targetForms map[string]bool) ([]UserFormCodeConflict, error) {
	if strings.TrimSpace(formsDir) == "" {
		return nil, nil
	}
	if _, err := os.Stat(formsDir); err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}
	sidecarDir := filepath.Join(formsDir, "code")
	conflicts := make([]UserFormCodeConflict, 0)
	err := filepath.WalkDir(formsDir, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			if samePath(path, sidecarDir) {
				return filepath.SkipDir
			}
			return nil
		}
		if !strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			return nil
		}
		formName := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if len(targetForms) > 0 && !targetForms[formName] {
			return nil
		}
		sidecarPath := filepath.Join(sidecarDir, formName+".bas")
		if _, err := os.Stat(sidecarPath); err != nil {
			if os.IsNotExist(err) {
				return nil
			}
			return err
		}
		formBody, err := os.ReadFile(path)
		if err != nil {
			return fmt.Errorf("read %s: %w", path, err)
		}
		sidecarBody, err := os.ReadFile(sidecarPath)
		if err != nil {
			return fmt.Errorf("read %s: %w", sidecarPath, err)
		}
		frmCode := NormalizeUserFormCodeText(ExtractUserFormCodeFromFRM(string(formBody)))
		sidecarCode := NormalizeUserFormCodeText(string(sidecarBody))
		if frmCode != sidecarCode {
			conflicts = append(conflicts, UserFormCodeConflict{
				FormName:    formName,
				FormPath:    path,
				SidecarPath: sidecarPath,
			})
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return conflicts, nil
}

func ExtractUserFormCodeFromFRM(text string) string {
	_, codeLines := splitUserFormFRMSections(text)
	if len(codeLines) == 0 {
		return ""
	}
	return strings.Join(codeLines, "\n")
}

func MergeUserFormCodeIntoFRM(frmText, sidecarText string) string {
	headerLines, _ := splitUserFormFRMSections(frmText)
	newline := detectTextNewline(frmText)
	code := NormalizeUserFormCodeText(sidecarText)
	out := append([]string{}, headerLines...)
	if code != "" {
		if len(out) > 0 {
			out = append(out, "")
		}
		out = append(out, splitNormalizedLines(code)...)
	}
	merged := strings.Join(out, newline)
	if merged == "" {
		return ""
	}
	return merged + newline
}

func splitUserFormFRMSections(text string) ([]string, []string) {
	if text == "" {
		return nil, nil
	}
	lines := splitNormalizedLines(text)
	start := 0
	lastAttribute := -1
	for i, line := range lines {
		if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
			lastAttribute = i
		}
	}
	if lastAttribute >= 0 {
		start = lastAttribute + 1
	}
	for start < len(lines) && strings.TrimSpace(lines[start]) == "" {
		start++
	}
	headerEnd := start
	if lastAttribute >= 0 {
		headerEnd = lastAttribute + 1
	}
	if headerEnd > len(lines) {
		headerEnd = len(lines)
	}
	header := append([]string{}, lines[:headerEnd]...)
	if start >= len(lines) {
		return header, nil
	}
	return header, append([]string{}, lines[start:]...)
}

func NormalizeUserFormCodeText(text string) string {
	normalized := strings.ReplaceAll(text, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	return strings.TrimRight(normalized, "\n")
}

func splitNormalizedLines(text string) []string {
	return strings.Split(strings.ReplaceAll(strings.ReplaceAll(text, "\r\n", "\n"), "\r", "\n"), "\n")
}

func detectTextNewline(text string) string {
	switch {
	case strings.Contains(text, "\r\n"):
		return "\r\n"
	case strings.Contains(text, "\r"):
		return "\r"
	default:
		return "\n"
	}
}

func samePath(a, b string) bool {
	return strings.EqualFold(filepath.Clean(a), filepath.Clean(b))
}
