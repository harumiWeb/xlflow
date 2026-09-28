package ovba

import (
	"fmt"
	"strings"
)

// ProjectComponentSpec describes one declaration to emit in the textual
// PROJECT stream.
type ProjectComponentSpec struct {
	Kind string
	Name string
}

type projectLine struct {
	body string
	eol  string
}

// RebuildProjectText replaces the component declarations in raw while
// preserving all unrelated PROJECT content byte-for-byte. Document and
// BaseClass declarations are template-owned and therefore must remain
// unchanged; Module and Class declarations may be added, removed, or renamed.
func RebuildProjectText(raw []byte, specs []ProjectComponentSpec) ([]byte, error) {
	lines := splitProjectLines(string(raw))
	original := ParseProjectText(raw).Components

	desired := make(map[string]ProjectComponentSpec, len(specs))
	for _, spec := range specs {
		if spec.Name == "" {
			return nil, fmt.Errorf("PROJECT component name is empty")
		}
		if !isProjectComponentKind(spec.Kind) {
			return nil, fmt.Errorf("PROJECT component %q has unknown kind %q", spec.Name, spec.Kind)
		}
		key := projectComponentKey(spec.Name)
		if prior, ok := desired[key]; ok {
			return nil, fmt.Errorf("duplicate PROJECT component names %q and %q", prior.Name, spec.Name)
		}
		desired[key] = spec
	}

	originalTemplateOwned := make(map[string]ProjectComponent)
	originalNames := make(map[string]ProjectComponent, len(original))
	for _, component := range original {
		key := projectComponentKey(component.Name)
		if prior, ok := originalNames[key]; ok {
			return nil, fmt.Errorf("duplicate PROJECT component declarations %s=%s and %s=%s", prior.Kind, prior.Name, component.Kind, component.Name)
		}
		originalNames[key] = component
		if component.Kind != "Document" && component.Kind != "BaseClass" {
			continue
		}
		originalTemplateOwned[key] = component
	}
	for _, component := range original {
		if component.Kind != "Document" && component.Kind != "BaseClass" {
			continue
		}
		key := projectComponentKey(component.Name)
		spec, ok := desired[key]
		if !ok || spec.Name != component.Name || spec.Kind != component.Kind {
			return nil, fmt.Errorf("template-owned PROJECT component %s=%s cannot be added, removed, or renamed", component.Kind, component.Name)
		}
	}
	for _, spec := range specs {
		if spec.Kind != "Document" && spec.Kind != "BaseClass" {
			continue
		}
		key := projectComponentKey(spec.Name)
		component, ok := originalTemplateOwned[key]
		if !ok || component.Name != spec.Name || component.Kind != spec.Kind {
			return nil, fmt.Errorf("template-owned PROJECT component %s=%s is not present in the template", spec.Kind, spec.Name)
		}
	}

	lastComponent := -1
	insertAfter := -1
	for i, line := range lines {
		_, ok := parseProjectComponentLine(line.body)
		if ok {
			lastComponent = i
			continue
		}
		if strings.HasPrefix(line.body, "ID=") {
			insertAfter = i
		}
	}
	if lastComponent >= 0 {
		insertAfter = lastComponent
	}

	var additions []ProjectComponentSpec
	for _, spec := range specs {
		if spec.Kind == "Document" || spec.Kind == "BaseClass" {
			continue
		}
		found := false
		for _, component := range original {
			if component.Name == spec.Name && component.Kind == spec.Kind {
				found = true
				break
			}
		}
		if !found {
			additions = append(additions, spec)
		}
	}

	eol := "\r\n"
	for _, line := range lines {
		if line.eol != "" {
			eol = line.eol
			break
		}
	}
	var out strings.Builder
	appendAdditions := func() {
		for _, spec := range additions {
			out.WriteString(spec.Kind)
			out.WriteByte('=')
			out.WriteString(spec.Name)
			out.WriteString(eol)
		}
	}
	if insertAfter < 0 {
		appendAdditions()
	}
	for i, line := range lines {
		component, isComponent := parseProjectComponentLine(line.body)
		keep := true
		if isComponent && component.Kind != "Document" && component.Kind != "BaseClass" {
			spec, ok := desired[projectComponentKey(component.Name)]
			keep = ok && spec.Name == component.Name && spec.Kind == component.Kind
		}
		if keep {
			out.WriteString(line.body)
			out.WriteString(line.eol)
		}
		if i == insertAfter {
			if keep && line.eol == "" && len(additions) > 0 {
				out.WriteString(eol)
			}
			appendAdditions()
		}
	}
	return []byte(out.String()), nil
}

func splitProjectLines(raw string) []projectLine {
	var lines []projectLine
	for len(raw) > 0 {
		i := strings.IndexByte(raw, '\n')
		if i < 0 {
			lines = append(lines, projectLine{body: raw})
			break
		}
		body, eol := raw[:i], "\n"
		if strings.HasSuffix(body, "\r") {
			body = strings.TrimSuffix(body, "\r")
			eol = "\r\n"
		}
		lines = append(lines, projectLine{body: body, eol: eol})
		raw = raw[i+1:]
	}
	return lines
}

func parseProjectComponentLine(line string) (ProjectComponent, bool) {
	key, val, ok := strings.Cut(line, "=")
	if !ok || !isProjectComponentKind(key) {
		return ProjectComponent{}, false
	}
	name := val
	if key == "Document" {
		name, _, _ = strings.Cut(name, "/")
	}
	name = strings.Trim(strings.TrimSpace(name), "\"")
	return ProjectComponent{Kind: key, Name: name}, true
}

func isProjectComponentKind(kind string) bool {
	switch kind {
	case "Module", "Class", "Document", "BaseClass":
		return true
	default:
		return false
	}
}

func projectComponentKey(name string) string {
	return strings.ToLower(name)
}
