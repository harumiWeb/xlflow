package doccomments

import (
	"fmt"
	"regexp"
	"sort"
	"strings"
)

type SymbolDocumentation struct {
	Summary          string            `json:"summary,omitempty"`
	Body             string            `json:"body,omitempty"`
	Parameters       map[string]string `json:"parameters,omitempty"`
	ParameterEntries []ParameterDoc    `json:"-"`
	Returns          string            `json:"returns,omitempty"`
	Errors           string            `json:"errors,omitempty"`
	Remarks          string            `json:"remarks,omitempty"`
	Examples         string            `json:"examples,omitempty"`
	SeeAlso          string            `json:"seeAlso,omitempty"`
	Deprecated       string            `json:"deprecated,omitempty"`
	UnknownSections  map[string]string `json:"unknownSections,omitempty"`
	Source           string            `json:"source,omitempty"`
}

type ParameterDoc struct {
	Name        string
	Description string
}

type Block struct {
	Doc        SymbolDocumentation
	StartLine  int
	EndLine    int
	StartByte  int
	EndByte    int
	LineTexts  []string
	Linked     bool
	Rubberduck string
}

type RubberduckAnnotation struct {
	Kind      string
	Text      string
	StartLine int
	EndLine   int
}

type Diagnostic struct {
	Code       string
	Message    string
	Line       int
	Column     int
	Suggestion string
}

type Parameter struct {
	Name     string
	Type     string
	Optional bool
}

type Procedure struct {
	Name       string
	Kind       string
	Parameters []Parameter
	ReturnType string
}

type Snippet struct {
	Label string
	Text  string
}

var (
	sectionHeaderRE = regexp.MustCompile(`^([A-Za-z][A-Za-z ]*):\s*$`)
	argEntryRE      = regexp.MustCompile(`^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(.*)$`)
	rubberduckRE    = regexp.MustCompile(`(?i)^\s*'\s*@([A-Za-z]*Description)\s*\(\s*"((?:[^"]|"")*)"\s*\)\s*$`)
)

func ParseDocLines(lines []string) SymbolDocumentation {
	contents := make([]string, 0, len(lines))
	for _, line := range lines {
		contents = append(contents, trimDocPrefix(line))
	}
	doc := SymbolDocumentation{
		Parameters:      map[string]string{},
		UnknownSections: map[string]string{},
		Source:          "doc_comment",
	}
	type section struct {
		name  string
		lines []string
	}
	sections := []section{{name: "summary"}}
	for _, content := range contents {
		if match := sectionHeaderRE.FindStringSubmatch(strings.TrimSpace(content)); match != nil {
			sections = append(sections, section{name: canonicalSection(match[1])})
			continue
		}
		idx := len(sections) - 1
		sections[idx].lines = append(sections[idx].lines, content)
	}
	for _, section := range sections {
		body := trimBlock(section.lines)
		switch section.name {
		case "summary":
			summary, rest := splitSummary(body)
			doc.Summary = summary
			doc.Body = rest
		case "args":
			for _, line := range section.lines {
				match := argEntryRE.FindStringSubmatch(line)
				if match == nil {
					continue
				}
				entry := ParameterDoc{Name: match[1], Description: strings.TrimSpace(match[2])}
				doc.ParameterEntries = append(doc.ParameterEntries, entry)
				doc.Parameters[entry.Name] = entry.Description
			}
		case "returns":
			doc.Returns = body
		case "errors":
			doc.Errors = body
		case "remarks":
			doc.Remarks = body
		case "examples":
			doc.Examples = body
		case "see also":
			doc.SeeAlso = body
		case "deprecated":
			doc.Deprecated = body
		default:
			if body != "" {
				doc.UnknownSections[section.name] = body
			}
		}
	}
	if len(doc.Parameters) == 0 {
		doc.Parameters = nil
	}
	if len(doc.UnknownSections) == 0 {
		doc.UnknownSections = nil
	}
	return doc
}

func RubberduckDescription(line string) (RubberduckAnnotation, bool) {
	match := rubberduckRE.FindStringSubmatch(line)
	if match == nil {
		return RubberduckAnnotation{}, false
	}
	return RubberduckAnnotation{Kind: normalizeRubberduckKind(match[1]), Text: strings.ReplaceAll(match[2], `""`, `"`)}, true
}

func RubberduckDocumentation(text string) SymbolDocumentation {
	return SymbolDocumentation{Summary: strings.TrimSpace(text), Source: "rubberduck"}
}

func DocumentationForTarget(source string, targetLine int, acceptedRubberduckKinds ...string) (SymbolDocumentation, int, bool) {
	lines := normalizedLines(source)
	if targetLine <= 1 || targetLine > len(lines)+1 {
		return SymbolDocumentation{}, 0, false
	}
	accepted := map[string]bool{}
	for _, kind := range acceptedRubberduckKinds {
		accepted[kind] = true
	}
	if len(accepted) == 0 {
		accepted["symbol"] = true
		accepted["variable"] = true
		accepted["module"] = true
	}
	rubberduck := ""
	for i := targetLine - 2; i >= 0; i-- {
		line := lines[i]
		if strings.TrimSpace(line) == "" {
			continue
		}
		if annotation, ok := RubberduckDescription(line); ok {
			if accepted[annotation.Kind] && rubberduck == "" {
				rubberduck = annotation.Text
			}
			continue
		}
		if isDocLine(line) {
			end := i
			start := i
			for start-1 >= 0 && isDocLine(lines[start-1]) {
				start--
			}
			doc := ParseDocLines(lines[start : end+1])
			return MergeDocAndRubberduck(doc, rubberduck), start + 1, true
		}
		break
	}
	if rubberduck != "" {
		return RubberduckDocumentation(rubberduck), targetLine - 1, true
	}
	return SymbolDocumentation{}, 0, false
}

func ModuleDocumentation(source string) (SymbolDocumentation, bool) {
	for _, line := range normalizedLines(source) {
		if annotation, ok := RubberduckDescription(line); ok && annotation.Kind == "module" {
			return RubberduckDocumentation(annotation.Text), true
		}
	}
	return SymbolDocumentation{}, false
}

func UnlinkedDocDiagnostics(source string, linkedStartLines map[int]bool, declarationLines map[int]bool) []Diagnostic {
	var out []Diagnostic
	lines := normalizedLines(source)
	for i := 0; i < len(lines); i++ {
		if !isDocLine(lines[i]) {
			continue
		}
		start := i
		for i+1 < len(lines) && isDocLine(lines[i+1]) {
			i++
		}
		startLine := start + 1
		if linkedStartLines[startLine] || docBlockCanLink(lines, i+1, declarationLines) {
			continue
		}
		out = append(out, Diagnostic{Code: "VB043", Message: "Documentation comment is not associated with a declaration.", Line: startLine, Column: 1})
	}
	return out
}

func MergeDocAndRubberduck(doc SymbolDocumentation, rubberduck string) SymbolDocumentation {
	if strings.TrimSpace(doc.Summary) == "" && strings.TrimSpace(rubberduck) != "" {
		doc.Summary = strings.TrimSpace(rubberduck)
	}
	if doc.Source == "" {
		doc.Source = "doc_comment"
	}
	return doc
}

func Validate(proc Procedure, doc SymbolDocumentation, startLine int) []Diagnostic {
	if !HasDocumentation(doc) {
		return nil
	}
	var out []Diagnostic
	seen := map[string]string{}
	entries := doc.ParameterEntries
	if len(entries) == 0 {
		for name, text := range doc.Parameters {
			entries = append(entries, ParameterDoc{Name: name, Description: text})
		}
	}
	for _, entry := range entries {
		name := entry.Name
		lower := strings.ToLower(name)
		if prior, ok := seen[lower]; ok {
			msg := fmt.Sprintf("Documentation parameter %q is listed more than once.", name)
			if prior != name {
				msg = fmt.Sprintf("Documentation parameters %q and %q differ only by case.", prior, name)
			}
			out = append(out, Diagnostic{Code: "VB041", Message: msg, Line: startLine, Column: 1})
			continue
		}
		seen[lower] = name
		if !hasParameter(proc.Parameters, name) {
			msg := fmt.Sprintf("Documentation references unknown parameter %q.", name)
			diag := Diagnostic{Code: "VB040", Message: msg, Line: startLine, Column: 1}
			if suggestion, ok := closestParameter(proc.Parameters, name); ok {
				diag.Message += fmt.Sprintf(" Did you mean %q?", suggestion)
				diag.Suggestion = suggestion
			}
			out = append(out, diag)
		}
	}
	if strings.TrimSpace(doc.Returns) != "" && strings.EqualFold(proc.Kind, "sub") {
		out = append(out, Diagnostic{Code: "VB042", Message: "Documentation has a Returns section, but Sub procedures do not return a value.", Line: startLine, Column: 1})
	}
	return out
}

func GenerateSnippet(proc Procedure) Snippet {
	if strings.TrimSpace(proc.Name) == "" || !snippetSupported(proc.Kind) {
		return Snippet{}
	}
	placeholder := "Summary."
	if strings.HasPrefix(strings.ToLower(proc.Kind), "property") {
		placeholder = "Property description."
	}
	index := 1
	var b strings.Builder
	fmt.Fprintf(&b, "''' ${%d:%s}", index, placeholder)
	index++
	args := snippetArgs(proc)
	if len(args) > 0 {
		b.WriteString("\n'''\n''' Args:")
		for _, param := range args {
			fmt.Fprintf(&b, "\n'''     %s: ${%d:%s}", param.Name, index, argPlaceholder(proc.Kind))
			index++
		}
	}
	if snippetNeedsReturns(proc) {
		b.WriteString("\n'''\n''' Returns:")
		fmt.Fprintf(&b, "\n'''     ${%d:%s}", index, returnPlaceholder(proc.Kind))
	}
	return Snippet{Label: "Generate documentation comment for " + proc.Name, Text: b.String()}
}

func HasDocumentation(doc SymbolDocumentation) bool {
	return strings.TrimSpace(doc.Summary) != "" ||
		strings.TrimSpace(doc.Body) != "" ||
		len(doc.Parameters) > 0 ||
		strings.TrimSpace(doc.Returns) != "" ||
		strings.TrimSpace(doc.Errors) != "" ||
		strings.TrimSpace(doc.Remarks) != "" ||
		strings.TrimSpace(doc.Examples) != "" ||
		strings.TrimSpace(doc.SeeAlso) != "" ||
		strings.TrimSpace(doc.Deprecated) != "" ||
		len(doc.UnknownSections) > 0
}

func Markdown(doc SymbolDocumentation, activeParameter string) string {
	if !HasDocumentation(doc) {
		return ""
	}
	if activeParameter != "" {
		for name, text := range doc.Parameters {
			if strings.EqualFold(name, activeParameter) && strings.TrimSpace(text) != "" {
				return strings.TrimSpace(text)
			}
		}
	}
	var b strings.Builder
	writeParagraph(&b, doc.Summary)
	writeParagraph(&b, doc.Body)
	if len(doc.Parameters) > 0 {
		writeHeading(&b, "Parameters")
		for _, entry := range SortedParameterDocs(doc) {
			writeParagraph(&b, "**"+entry.Name+"**\n\n"+strings.TrimSpace(entry.Description))
		}
	}
	writeSection(&b, "Returns", doc.Returns)
	writeSection(&b, "Errors", doc.Errors)
	writeSection(&b, "Remarks", doc.Remarks)
	writeSection(&b, "Examples", doc.Examples)
	writeSection(&b, "See Also", doc.SeeAlso)
	writeSection(&b, "Deprecated", doc.Deprecated)
	for name, text := range doc.UnknownSections {
		writeSection(&b, name, text)
	}
	return strings.TrimSpace(b.String())
}

func SortedParameterDocs(doc SymbolDocumentation) []ParameterDoc {
	if len(doc.ParameterEntries) > 0 {
		out := make([]ParameterDoc, len(doc.ParameterEntries))
		copy(out, doc.ParameterEntries)
		return out
	}
	out := make([]ParameterDoc, 0, len(doc.Parameters))
	for name, text := range doc.Parameters {
		out = append(out, ParameterDoc{Name: name, Description: text})
	}
	sort.SliceStable(out, func(i, j int) bool { return strings.ToLower(out[i].Name) < strings.ToLower(out[j].Name) })
	return out
}

func trimDocPrefix(line string) string {
	trimmed := strings.TrimLeft(line, " \t")
	if strings.HasPrefix(trimmed, "'''") {
		return strings.TrimPrefix(trimmed, "'''")
	}
	return line
}

func isDocLine(line string) bool {
	return strings.HasPrefix(strings.TrimLeft(line, " \t"), "'''")
}

func normalizedLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}

func docBlockCanLink(lines []string, afterBlock int, declarationLines map[int]bool) bool {
	for i := afterBlock; i < len(lines); i++ {
		if strings.TrimSpace(lines[i]) == "" {
			continue
		}
		if _, ok := RubberduckDescription(lines[i]); ok {
			continue
		}
		return declarationLines[i+1]
	}
	return false
}

func canonicalSection(name string) string {
	return strings.ToLower(strings.Join(strings.Fields(strings.TrimSpace(name)), " "))
}

func trimBlock(lines []string) string {
	start, end := 0, len(lines)
	for start < end && strings.TrimSpace(lines[start]) == "" {
		start++
	}
	for end > start && strings.TrimSpace(lines[end-1]) == "" {
		end--
	}
	out := make([]string, 0, end-start)
	for _, line := range lines[start:end] {
		out = append(out, strings.TrimRight(line, " \t"))
	}
	return strings.TrimSpace(strings.Join(out, "\n"))
}

func splitSummary(text string) (string, string) {
	parts := strings.Split(text, "\n")
	var first []string
	i := 0
	for ; i < len(parts); i++ {
		if strings.TrimSpace(parts[i]) == "" {
			break
		}
		first = append(first, strings.TrimSpace(parts[i]))
	}
	summary := strings.TrimSpace(strings.Join(first, "\n"))
	rest := strings.TrimSpace(strings.Join(parts[min(i+1, len(parts)):], "\n"))
	return summary, rest
}

func normalizeRubberduckKind(kind string) string {
	switch strings.ToLower(kind) {
	case "moduledescription":
		return "module"
	case "variabledescription":
		return "variable"
	default:
		return "symbol"
	}
}

func hasParameter(params []Parameter, name string) bool {
	for _, param := range params {
		if strings.EqualFold(param.Name, name) {
			return true
		}
	}
	return false
}

func closestParameter(params []Parameter, name string) (string, bool) {
	candidates := make([]string, 0, len(params))
	for _, param := range params {
		candidates = append(candidates, param.Name)
	}
	return closestName(candidates, name)
}

func closestName(candidates []string, name string) (string, bool) {
	best := ""
	bestDistance := 4
	for _, candidate := range candidates {
		if candidate == "" || strings.EqualFold(candidate, name) {
			continue
		}
		distance := editDistance(strings.ToLower(name), strings.ToLower(candidate))
		if distance < bestDistance {
			best = candidate
			bestDistance = distance
		}
	}
	limit := 1
	if len(name) >= 4 {
		limit = 2
	}
	return best, best != "" && bestDistance <= limit
}

func editDistance(a, b string) int {
	ar, br := []rune(a), []rune(b)
	prev := make([]int, len(br)+1)
	for j := range prev {
		prev[j] = j
	}
	for i, ra := range ar {
		current := make([]int, len(br)+1)
		current[0] = i + 1
		for j, rb := range br {
			cost := 0
			if ra != rb {
				cost = 1
			}
			current[j+1] = min(min(current[j]+1, prev[j+1]+1), prev[j]+cost)
		}
		prev = current
	}
	return prev[len(br)]
}

func snippetSupported(kind string) bool {
	switch strings.ToLower(kind) {
	case "sub", "function", "property_get", "property_let", "property_set":
		return true
	default:
		return false
	}
}

func snippetArgs(proc Procedure) []Parameter {
	if strings.EqualFold(proc.Kind, "property_get") {
		return nil
	}
	return proc.Parameters
}

func snippetNeedsReturns(proc Procedure) bool {
	return strings.EqualFold(proc.Kind, "function") || strings.EqualFold(proc.Kind, "property_get")
}

func argPlaceholder(kind string) string {
	if strings.EqualFold(kind, "property_let") || strings.EqualFold(kind, "property_set") {
		return "Value description."
	}
	return "Parameter description."
}

func returnPlaceholder(kind string) string {
	if strings.EqualFold(kind, "property_get") {
		return "Returned property value description."
	}
	return "Return value description."
}

func writeHeading(b *strings.Builder, heading string) {
	if b.Len() > 0 {
		b.WriteString("\n\n")
	}
	b.WriteString("**")
	b.WriteString(heading)
	b.WriteString("**")
}

func writeSection(b *strings.Builder, heading, text string) {
	if strings.TrimSpace(text) == "" {
		return
	}
	writeHeading(b, heading)
	writeParagraph(b, text)
}

func writeParagraph(b *strings.Builder, text string) {
	text = strings.TrimSpace(text)
	if text == "" {
		return
	}
	if b.Len() > 0 {
		b.WriteString("\n\n")
	}
	b.WriteString(text)
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
