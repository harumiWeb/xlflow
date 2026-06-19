package userforms

import (
	"regexp"
	"strings"
)

type Form struct {
	Name     string
	Controls []Control
}

type Control struct {
	Name        string
	Type        string
	StartLine   int
	StartColumn int
	EndLine     int
	EndColumn   int
}

var beginRe = regexp.MustCompile(`(?i)^\s*Begin\s+(\S+)\s+([A-Za-z_][A-Za-z0-9_]*)\b`)
var attrNameRe = regexp.MustCompile(`(?i)^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"`)

func Parse(source string) Form {
	lines := normalizedLines(source)
	form := Form{}
	type stackEntry struct {
		name        string
		typ         string
		startLine   int
		startColumn int
	}
	var stack []stackEntry
	for i, line := range lines {
		if m := beginRe.FindStringSubmatch(line); len(m) == 3 {
			name := m[2]
			typ := normalizeType(m[1])
			col := strings.Index(line, name)
			if col < 0 {
				col = 0
			}
			stack = append(stack, stackEntry{name: name, typ: typ, startLine: i + 1, startColumn: col + 1})
			if len(stack) == 1 && form.Name == "" {
				form.Name = name
			}
			continue
		}
		if m := attrNameRe.FindStringSubmatch(line); len(m) == 2 && form.Name == "" {
			form.Name = m[1]
			continue
		}
		if strings.EqualFold(strings.TrimSpace(line), "End") && len(stack) > 0 {
			entry := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			if len(stack) >= 1 && entry.name != "" && entry.typ != "" {
				form.Controls = append(form.Controls, Control{
					Name:        entry.name,
					Type:        entry.typ,
					StartLine:   entry.startLine,
					StartColumn: entry.startColumn,
					EndLine:     i + 1,
					EndColumn:   len(line) + 1,
				})
			}
		}
	}
	return form
}

func normalizeType(raw string) string {
	raw = strings.TrimSpace(raw)
	if raw == "" || strings.HasPrefix(raw, "{") {
		return ""
	}
	parts := strings.Split(raw, ".")
	for len(parts) > 0 && isNumeric(parts[len(parts)-1]) {
		parts = parts[:len(parts)-1]
	}
	if len(parts) == 0 {
		return raw
	}
	name := parts[len(parts)-1]
	switch strings.ToLower(parts[0]) {
	case "forms", "msforms":
		return "MSForms." + name
	default:
		return raw
	}
}

func isNumeric(s string) bool {
	if s == "" {
		return false
	}
	for _, r := range s {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func normalizedLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}
