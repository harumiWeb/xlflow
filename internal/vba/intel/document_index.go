package intel

import (
	"regexp"
	"strings"
	"sync"
)

var (
	setAssignmentIndexRe = regexp.MustCompile(`(?im)\bSet\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^\r\n:]+)`)
	withStartIndexRe     = regexp.MustCompile(`(?i)^With\s+(.+)$`)
	withEndIndexRe       = regexp.MustCompile(`(?i)^End\s+With\b`)
	newAssignmentExprRe  = regexp.MustCompile(`(?i)^New\s+([A-Za-z_][A-Za-z0-9_.]*)\b`)
	createObjectExprRe   = regexp.MustCompile(`(?i)^CreateObject\s*\(\s*"([^"]+)"\s*\)`)
)

// documentIndex lazily initializes each derived lookup at most once and never
// mutates it afterward. All lookup keys are folded because VBA identifiers are
// case-insensitive.
type documentIndex struct {
	source         string
	procedures     []ProcedureInfo
	procedureLines []int

	symbolsByName map[string][]Symbol
	withBlocks    []withBlockRange
	withByLine    []int

	assignmentsOnce   sync.Once
	assignmentsByName map[string][]typedAssignment
}

type typedAssignment struct {
	name       string
	procedure  string
	position   Position
	offset     int
	expression string
}

type withBlockRange struct {
	receiver string
	parent   int
	range_   Range
}

func indexName(name string) string { return strings.ToLower(strings.TrimSpace(name)) }

func buildDocumentIndex(source string, lines []string, procedures []ProcedureInfo, procedureLines []int, symbols []Symbol) *documentIndex {
	idx := &documentIndex{
		source:         source,
		procedures:     append([]ProcedureInfo(nil), procedures...),
		procedureLines: append([]int(nil), procedureLines...),
		symbolsByName:  make(map[string][]Symbol),
		withByLine:     make([]int, len(lines)),
	}
	for i := range idx.withByLine {
		idx.withByLine[i] = -1
	}
	for _, sym := range symbols {
		idx.addSymbol(sym)
		if procedureReturnSymbol(sym) {
			idx.addSymbol(Symbol{
				Name:       sym.Name,
				Kind:       "function_return",
				ReturnType: sym.ReturnType,
				File:       sym.File,
				Parent:     sym.Name,
				Range:      sym.Range,
				Selection:  sym.Selection,
			})
		}
	}
	idx.buildWithBlocks(lines)
	return idx
}

func (idx *documentIndex) initAssignments() {
	if idx == nil {
		return
	}
	idx.assignmentsOnce.Do(func() {
		idx.assignmentsByName = make(map[string][]typedAssignment)
		for _, match := range setAssignmentIndexRe.FindAllStringSubmatchIndex(idx.source, -1) {
			if len(match) < 6 || match[2] < 0 || match[3] < 0 || match[4] < 0 || match[5] < 0 {
				continue
			}
			raw := stripLineComment(idx.source[match[4]:match[5]])
			expr := strings.TrimSpace(raw)
			if expr == "" {
				continue
			}
			exprOffset := match[4] + strings.Index(raw, expr)
			name := idx.source[match[2]:match[3]]
			position := positionForByteOffset(idx.source, match[0])
			idx.assignmentsByName[indexName(name)] = append(idx.assignmentsByName[indexName(name)], typedAssignment{
				name:       name,
				procedure:  procedureNameAt(idx.procedures, idx.procedureLines, position),
				position:   position,
				offset:     exprOffset,
				expression: expr,
			})
		}
	})
}

func (idx *documentIndex) addSymbol(sym Symbol) {
	if idx == nil || indexName(sym.Name) == "" {
		return
	}
	key := indexName(sym.Name)
	idx.symbolsByName[key] = append(idx.symbolsByName[key], sym)
}

func procedureReturnSymbol(sym Symbol) bool {
	return sym.ReturnType != "" && (strings.EqualFold(sym.Kind, "function") || strings.EqualFold(sym.Kind, "property_get"))
}

func (idx *documentIndex) buildWithBlocks(lines []string) {
	stack := make([]int, 0)
	for lineNo, line := range lines {
		if len(stack) > 0 {
			idx.withByLine[lineNo] = stack[len(stack)-1]
		}
		text := strings.TrimSpace(stripLineComment(line))
		if text == "" {
			continue
		}
		if withEndIndexRe.MatchString(text) {
			if len(stack) > 0 {
				block := stack[len(stack)-1]
				idx.withBlocks[block].range_.End = Position{Line: lineNo, Character: utf16Len(line)}
				stack = stack[:len(stack)-1]
			}
			continue
		}
		match := withStartIndexRe.FindStringSubmatch(text)
		if len(match) == 0 {
			continue
		}
		parent := -1
		if len(stack) > 0 {
			parent = stack[len(stack)-1]
		}
		idx.withBlocks = append(idx.withBlocks, withBlockRange{
			receiver: strings.TrimSpace(match[1]), parent: parent,
			range_: Range{Start: Position{Line: lineNo}, End: Position{Line: len(lines)}},
		})
		stack = append(stack, len(idx.withBlocks)-1)
	}
}

func (idx *documentIndex) withBlockAt(pos Position) (int, bool) {
	if idx == nil || pos.Line < 0 || pos.Line >= len(idx.withByLine) {
		return 0, false
	}
	block := idx.withByLine[pos.Line]
	if block < 0 || block >= len(idx.withBlocks) {
		return 0, false
	}
	return block, true
}

func (idx *documentIndex) nearestAssignment(name, procedure string, pos Position) (typedAssignment, bool) {
	if idx == nil {
		return typedAssignment{}, false
	}
	idx.initAssignments()
	assignments := idx.assignmentsByName[indexName(name)]
	for i := len(assignments) - 1; i >= 0; i-- {
		if comparePosition(assignments[i].position, pos) <= 0 && (procedure == "" || assignments[i].procedure == "" || strings.EqualFold(assignments[i].procedure, procedure)) {
			return assignments[i], true
		}
	}
	return typedAssignment{}, false
}

func procedureNameAt(procedures []ProcedureInfo, procedureLines []int, pos Position) string {
	if pos.Line < 0 || pos.Line >= len(procedureLines) {
		return ""
	}
	index := procedureLines[pos.Line]
	if index < 0 || index >= len(procedures) {
		return ""
	}
	return procedures[index].Name
}

func procedureIndexForLines(lines []string) ([]ProcedureInfo, []int) {
	procedureLines := make([]int, len(lines))
	for i := range procedureLines {
		procedureLines[i] = -1
	}
	procedures := make([]ProcedureInfo, 0)
	depth, active := 0, -1
	for lineNo, line := range lines {
		text := strings.TrimSpace(line[:codeLimit(line)])
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		switch {
		case strings.HasPrefix(lower, "end sub") || strings.HasPrefix(lower, "end function") || strings.HasPrefix(lower, "end property"):
			if depth > 0 {
				depth--
			}
			if depth == 0 && active >= 0 {
				procedures[active].Range.End = Position{Line: lineNo, Character: utf16Len(line)}
				active = -1
			}
		case procedureStartLine(lower):
			depth++
			if depth == 1 {
				if name := procedureNameFromLine(text); name != "" {
					procedures = append(procedures, ProcedureInfo{Name: name, Range: Range{Start: Position{Line: lineNo}, End: Position{Line: len(lines)}}})
					active = len(procedures) - 1
				}
			}
		}
	}
	for index, procedure := range procedures {
		lastLine := min(procedure.Range.End.Line, len(procedureLines)-1)
		for lineNo := procedure.Range.Start.Line; lineNo <= lastLine && lineNo >= 0; lineNo++ {
			procedureLines[lineNo] = index
		}
	}
	return procedures, procedureLines
}
