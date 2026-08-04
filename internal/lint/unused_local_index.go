package lint

import (
	"context"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// unusedLocalReferenceIndexStats exposes structural work to package tests. It
// deliberately counts procedure scans rather than wall-clock time so the
// regression is stable across machines.
type unusedLocalReferenceIndexStats struct {
	SourceNormalizations int
	ProcedureScans       int
	StatementsVisited    int
	IdentifiersVisited   int
}

type localReferenceProcedure struct {
	name      string
	startLine int
	endLine   int
	locals    map[string][]int
}

// buildProcedureLocalReferenceIndexContext scans each procedure once and
// records read occurrences for every local declared in that procedure. The
// returned slice is aligned with allSymbols.
func buildProcedureLocalReferenceIndexContext(ctx context.Context, source string, allSymbols []symbols.Symbol, stats *unusedLocalReferenceIndexStats) ([]bool, error) {
	lines := normalizedSourceLines(source)
	if stats != nil {
		stats.SourceNormalizations++
	}

	procedures := make([]localReferenceProcedure, 0)
	for symbolIndex, sym := range allSymbols {
		if symbolIndex&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		if !procedureSymbolKind(sym.Kind) {
			continue
		}
		procedures = append(procedures, localReferenceProcedure{
			name:      sym.Name,
			startLine: sym.StartLine,
			endLine:   sym.EndLine,
			locals:    make(map[string][]int),
		})
	}

	// Associate locals by parent and containment. Property accessors can share a
	// name, so the source range is required to select the right one.
	for symbolIndex, sym := range allSymbols {
		if symbolIndex&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		if !isUnusedLocalCandidate(sym) || isIgnoredLocalName(sym.Name) {
			continue
		}
		procedureIndex := containingProcedureIndex(procedures, sym)
		if procedureIndex < 0 {
			// Preserve the former conservative fallback for recovered symbols by
			// indexing only the range carried by a parented local. An unparented
			// recovered local retains the former scan-to-EOF behavior.
			endLine := sym.EndLine
			if sym.Parent == "" {
				endLine = len(lines)
			}
			procedures = append(procedures, localReferenceProcedure{
				startLine: sym.StartLine,
				endLine:   endLine,
				locals:    make(map[string][]int),
			})
			procedureIndex = len(procedures) - 1
		}
		name := strings.ToLower(sym.Name)
		procedures[procedureIndex].locals[name] = append(procedures[procedureIndex].locals[name], symbolIndex)
	}

	referenced := make([]bool, len(allSymbols))
	for procedureIndex := range procedures {
		procedure := &procedures[procedureIndex]
		if len(procedure.locals) == 0 {
			continue
		}
		if stats != nil {
			stats.ProcedureScans++
		}
		start := max(procedure.startLine, 1)
		end := min(procedure.endLine, len(lines))
		for lineNo := start; lineNo <= end; lineNo++ {
			if lineNo&0xff == 0 {
				if err := ctx.Err(); err != nil {
					return nil, err
				}
			}
			line := normalizedCodeLine(lines[lineNo-1])
			for _, statement := range splitStatements(line) {
				if stats != nil {
					stats.StatementsVisited++
				}
				indexLocalReads(statement, lineNo, procedure.locals, allSymbols, referenced, stats)
			}
		}
	}
	return referenced, ctx.Err()
}

func isUnusedLocalCandidate(sym symbols.Symbol) bool {
	return sym.Kind == "local_variable" || (sym.Kind == "const" && sym.Parent != "")
}

func containingProcedureIndex(procedures []localReferenceProcedure, local symbols.Symbol) int {
	if local.Parent == "" {
		return -1
	}
	for i := range procedures {
		procedure := procedures[i]
		if !strings.EqualFold(procedure.name, local.Parent) {
			continue
		}
		if procedure.startLine > local.StartLine || local.StartLine > procedure.endLine {
			continue
		}
		return i
	}
	return -1
}

func indexLocalReads(statement string, lineNo int, locals map[string][]int, allSymbols []symbols.Symbol, referenced []bool, stats *unusedLocalReferenceIndexStats) {
	declaration := isLocalDeclarationStatement(statement)
	for start := 0; start < len(statement); {
		if !isLocalReferenceIdentifierByte(statement[start]) {
			start++
			continue
		}
		end := start + 1
		for end < len(statement) && isLocalReferenceIdentifierByte(statement[end]) {
			end++
		}
		if stats != nil {
			stats.IdentifiersVisited++
		}
		indices := locals[strings.ToLower(statement[start:end])]
		if len(indices) != 0 &&
			!isQualifiedIdentifier(statement, start) &&
			!isLocalAssignmentTargetOccurrence(statement, start, end) {
			for _, symbolIndex := range indices {
				sym := allSymbols[symbolIndex]
				if lineNo < sym.StartLine || (declaration && lineNo == sym.StartLine) {
					continue
				}
				referenced[symbolIndex] = true
			}
		}
		start = end
	}
}

// VBA identifiers can contain non-ASCII letters. Treating a UTF-8 sequence as
// one token preserves those names while the existing ASCII identifier rules
// continue to define qualification and assignment-target boundaries.
func isLocalReferenceIdentifierByte(b byte) bool {
	return isIdentifierChar(b) || b >= 0x80
}
