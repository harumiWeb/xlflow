package ovba

import (
	"bytes"
	"fmt"
	"strings"
)

// ProjectComponentSpec describes one declaration to emit in the textual
// PROJECT stream.
type ProjectComponentSpec struct {
	Kind string
	Name string
}

// projectLine is one physical line of the PROJECT stream. body and eol keep
// the raw bytes so unchanged lines can be re-emitted byte-for-byte.
type projectLine struct {
	body []byte
	eol  []byte
}

// RebuildProjectText replaces the component declarations in raw while
// preserving all unrelated PROJECT content byte-for-byte. Document and
// BaseClass declarations are template-owned and therefore must remain
// unchanged; Module and Class declarations may be added, removed, or renamed.
// The stream text is MBCS in the project code page: component names are
// decoded for comparison and added lines are encoded back with the same
// code page.
func RebuildProjectText(raw []byte, specs []ProjectComponentSpec, codepage uint16) ([]byte, error) {
	lines := splitProjectLines(raw)
	parsed, err := ParseProjectText(raw, codepage)
	if err != nil {
		return nil, err
	}
	original := parsed.Components

	desired := make(map[string]ProjectComponentSpec, len(specs))
	for _, spec := range specs {
		if spec.Name == "" {
			return nil, fmt.Errorf("PROJECT component name is empty")
		}
		if !isProjectComponentKind(spec.Kind) {
			return nil, fmt.Errorf("PROJECT component %q has unknown kind %q", spec.Name, spec.Kind)
		}
		// A name is written raw into a textual "Kind=Name" line and re-parsed
		// with quotes and surrounding whitespace stripped, so anything that
		// could inject or corrupt a line is rejected here regardless of what
		// the caller validated.
		if strings.ContainsAny(spec.Name, "\x00\r\n\"") || strings.TrimSpace(spec.Name) != spec.Name {
			return nil, fmt.Errorf("PROJECT component %q contains characters a declaration line cannot carry", spec.Name)
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
		_, ok, err := parseProjectComponentLine(line.body, codepage)
		if err != nil {
			return nil, err
		}
		if ok {
			lastComponent = i
			continue
		}
		if bytes.HasPrefix(line.body, []byte("ID=")) {
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
	// Added names must be representable in the project code page; validate
	// before emitting so a failure cannot leave a partially encoded line.
	for _, spec := range additions {
		if _, err := EncodeMBCS(spec.Name, codepage); err != nil {
			return nil, fmt.Errorf("PROJECT component %q: %w", spec.Name, err)
		}
	}

	eol := []byte("\r\n")
	for _, line := range lines {
		if len(line.eol) != 0 {
			eol = line.eol
			break
		}
	}
	var out bytes.Buffer
	appendAdditions := func() {
		for _, spec := range additions {
			enc, _ := EncodeMBCS(spec.Name, codepage) // pre-validated above
			out.WriteString(spec.Kind)
			out.WriteByte('=')
			out.Write(enc)
			out.Write(eol)
		}
	}
	if insertAfter < 0 {
		appendAdditions()
	}
	for i, line := range lines {
		component, isComponent, err := parseProjectComponentLine(line.body, codepage)
		if err != nil {
			return nil, err
		}
		keep := true
		if isComponent && component.Kind != "Document" && component.Kind != "BaseClass" {
			spec, ok := desired[projectComponentKey(component.Name)]
			keep = ok && spec.Name == component.Name && spec.Kind == component.Kind
		}
		if keep {
			out.Write(line.body)
			out.Write(line.eol)
		}
		if i == insertAfter {
			if keep && len(line.eol) == 0 && len(additions) > 0 {
				out.Write(eol)
			}
			appendAdditions()
		}
	}
	return out.Bytes(), nil
}

func splitProjectLines(raw []byte) []projectLine {
	var lines []projectLine
	for len(raw) > 0 {
		i := bytes.IndexByte(raw, '\n')
		if i < 0 {
			lines = append(lines, projectLine{body: raw})
			break
		}
		body, eol := raw[:i], []byte("\n")
		if bytes.HasSuffix(body, []byte("\r")) {
			body = body[:len(body)-1]
			eol = []byte("\r\n")
		}
		lines = append(lines, projectLine{body: body, eol: eol})
		raw = raw[i+1:]
	}
	return lines
}

func parseProjectComponentLine(line []byte, codepage uint16) (ProjectComponent, bool, error) {
	key, val, ok := bytes.Cut(line, []byte("="))
	if !ok || !isProjectComponentKindBytes(key) {
		return ProjectComponent{}, false, nil
	}
	name := val
	if string(key) == "Document" {
		name, _, _ = bytes.Cut(name, []byte("/"))
	}
	name = bytes.Trim(bytes.TrimSpace(name), "\"")
	decoded, err := DecodeMBCS(name, codepage)
	if err != nil {
		return ProjectComponent{}, false, fmt.Errorf("PROJECT component %s: %w", key, err)
	}
	return ProjectComponent{Kind: string(key), Name: decoded}, true, nil
}

func isProjectComponentKind(kind string) bool {
	switch kind {
	case "Module", "Class", "Document", "BaseClass":
		return true
	default:
		return false
	}
}

func isProjectComponentKindBytes(kind []byte) bool {
	return isProjectComponentKind(string(kind))
}

func projectComponentKey(name string) string {
	return strings.ToLower(name)
}
