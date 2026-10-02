package filepush

import (
	"path/filepath"
	"regexp"
	"strings"
)

// Folder annotation helpers ported from VbaSourceHelper so the file backend
// writes '@Folder("A.B")' comments exactly where the Excel push writes them:
// the annotation replaces an existing annotation line in place, or is inserted
// before the first content line that is not a VERSION/Attribute header and not
// blank. Updated text is joined with CRLF, matching the .NET Environment
// newline used on the bridge host.
var (
	folderAnnotationPattern = regexp.MustCompile(`(?i)^'?@Folder\(\s*"([^"]*)"\s*\)`)
	attributeVBPattern      = regexp.MustCompile(`(?i)^Attribute\s+VB_`)
)

// updateFolderAnnotationText ports UpdateFolderAnnotationText. Modes "ignore"
// and "preserve" leave the text untouched; any other value is treated as the
// "update" contract validated by config.
func updateFolderAnnotationText(text, mode, desiredAnnotation string) string {
	if mode == "ignore" || mode == "preserve" {
		return text
	}
	lines := splitLinesKeepEmpty(text)
	found, lineIndex := findFolderAnnotationLine(lines)
	annotationLine := ""
	if strings.TrimSpace(desiredAnnotation) != "" {
		annotationLine = `'@Folder("` + desiredAnnotation + `")`
	}
	switch {
	case found && annotationLine == "":
		lines = append(lines[:lineIndex], lines[lineIndex+1:]...)
	case found:
		lines[lineIndex] = annotationLine
	case annotationLine != "":
		insertIndex := folderAnnotationInsertIndex(lines)
		lines = append(lines[:insertIndex], append([]string{annotationLine}, lines[insertIndex:]...)...)
	}
	return strings.Join(lines, "\r\n")
}

// relativePathSegments ports GetRelativePathSegments: the file's parent
// directory relative to its source root, split into cleaned segments. Files
// directly under the root (or outside it, which cannot happen for discovered
// sources) produce no segments.
func relativePathSegments(rootDir, filePath string) []string {
	if strings.TrimSpace(rootDir) == "" || strings.TrimSpace(filePath) == "" {
		return nil
	}
	rootFull, err := filepath.Abs(rootDir)
	if err != nil {
		return nil
	}
	fileFull, err := filepath.Abs(filePath)
	if err != nil {
		return nil
	}
	rootFull = filepath.Clean(rootFull)
	parentDir := filepath.Dir(filepath.Clean(fileFull))
	if !strings.HasPrefix(strings.ToLower(parentDir), strings.ToLower(rootFull)) {
		return nil
	}
	relative, err := filepath.Rel(rootFull, parentDir)
	if err != nil || relative == "" || relative == "." {
		return nil
	}
	var segments []string
	for _, part := range strings.FieldsFunc(relative, func(r rune) bool { return r == '/' || r == '\\' }) {
		if part == ".." {
			// A path resolving outside its root is a layout error, not an
			// annotation opportunity; surface it upstream instead of silently
			// producing a bogus annotation.
			return nil
		}
		if cleaned := cleanFolderPathSegment(part); cleaned != "" {
			segments = append(segments, cleaned)
		}
	}
	return segments
}

// buildFolderAnnotation ports BuildFolderAnnotation: cleaned segments joined
// with dots, or "" when nothing usable remains.
func buildFolderAnnotation(segments []string) string {
	clean := make([]string, 0, len(segments))
	for _, segment := range segments {
		if cleaned := cleanFolderPathSegment(segment); cleaned != "" {
			clean = append(clean, cleaned)
		}
	}
	return strings.Join(clean, ".")
}

func folderAnnotationForPath(rootDir, filePath string) string {
	return buildFolderAnnotation(relativePathSegments(rootDir, filePath))
}

func findFolderAnnotationLine(lines []string) (bool, int) {
	for i, line := range lines {
		if folderAnnotationPattern.MatchString(strings.TrimSpace(line)) {
			return true, i
		}
	}
	return false, -1
}

func folderAnnotationInsertIndex(lines []string) int {
	for i, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "VERSION 1.0 CLASS" || attributeVBPattern.MatchString(trimmed) {
			continue
		}
		if trimmed == "" {
			continue
		}
		return i
	}
	return 0
}

// cleanFolderPathSegment ports CleanFolderPathSegment: trim, then replace
// characters that are invalid in Windows file names with '_'. The invalid set
// is fixed rather than platform-dependent so Linux/WSL runs produce the same
// annotation text the Windows bridge produces.
func cleanFolderPathSegment(segment string) string {
	cleaned := strings.TrimSpace(segment)
	if cleaned == "" {
		return ""
	}
	var b strings.Builder
	for _, r := range cleaned {
		if r < 32 || strings.ContainsRune(`"<>|?*:\/`, r) {
			b.WriteByte('_')
			continue
		}
		b.WriteRune(r)
	}
	return b.String()
}

// splitLinesKeepEmpty ports VbaSourceHelper.SplitLines: split on every line
// break style, keeping empty elements so line positions stay stable.
func splitLinesKeepEmpty(text string) []string {
	if text == "" {
		return nil
	}
	normalized := strings.ReplaceAll(text, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	return strings.Split(normalized, "\n")
}
