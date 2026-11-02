package analyze

import (
	"fmt"
	"regexp"
	"slices"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestUnusedPrivateConstLexicalIndexPreservesMatchLocations(t *testing.T) {
	t.Parallel()
	names := []string{"A", "Limit", "K", "S", "変数"}
	for _, boundaries := range [][]string{
		{"", ".", "X", "_", "é", "\xff", " ", "A."},
		{"", ".", "K", "ſ", "\xff"},
	} {
		var lines []string
		for _, name := range names {
			for _, prefix := range boundaries {
				for _, suffix := range boundaries {
					lines = append(lines, prefix+name+suffix+name+prefix)
				}
			}
		}
		index := buildUnusedPrivateConstLexicalIndex(lines, names)
		for _, name := range names {
			var expected []unusedPrivateConstOccurrence
			matcher := unusedNameMatcher(name)
			for lineNumber, line := range lines {
				for _, location := range matcher.FindAllStringIndex(line, -1) {
					begin := location[0]
					if !isIdentifierByte(line[begin]) {
						begin++
					}
					expected = append(expected, unusedPrivateConstOccurrence{
						line: lineNumber + 1, begin: begin, matchFrom: location[0], matchTo: location[1],
					})
				}
			}
			if !slices.Equal(index.byName[name], expected) {
				got := index.byName[name]
				for i := range min(len(got), len(expected)) {
					if got[i] != expected[i] {
						t.Fatalf("%s occurrence %d = %+v, regexp %+v; source=%q", name, i, got[i], expected[i], lines[expected[i].line-1])
					}
				}
				t.Fatalf("%s indexed locations count %d differs from regexp matches %d", name, len(got), len(expected))
			}
		}
	}
}

func TestUnusedPrivateConstLexicalIndexMatchesRegexpSemantics(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name           string
		candidate      string
		lines          []string
		declarator     vbaast.Range
		procedureRange []vbaast.Range
		shadowed       map[int]bool
		moduleName     string
		wantReferenced bool
	}{
		{
			name:      "comments strings directives case and boundaries",
			candidate: "Limit",
			lines: []string{
				"' Limit in a comment",
				"Debug.Print \"\u00e9LIMIT\u00e9\"",
				"#If Limit Then",
				"XLimit Limitations Limit_",
			},
			wantReferenced: true,
		},
		{
			name:           "regexp nonoverlap keeps own declarator excluded",
			candidate:      "A",
			lines:          []string{"A.A"},
			declarator:     vbaast.Range{StartLine: 1, StartColumn: 1, EndLine: 1, EndColumn: 2},
			wantReferenced: false,
		},
		{
			name:           "Unicode simple-fold fallback",
			candidate:      "K",
			lines:          []string{"Debug.Print \"\u212a\""},
			wantReferenced: true,
		},
		{
			name:           "Unicode long-s fallback",
			candidate:      "S",
			lines:          []string{"Debug.Print \"\u017f\""},
			wantReferenced: true,
		},
		{
			name:           "same-name local shadow suppresses unqualified use",
			candidate:      "Limit",
			lines:          []string{"Public Sub Run()", "  Limit = 1", "End Sub"},
			procedureRange: []vbaast.Range{{StartLine: 1, StartColumn: 1, EndLine: 3, EndColumn: 1, StartByte: 41}},
			shadowed:       map[int]bool{41: true},
			moduleName:     "Main",
			wantReferenced: false,
		},
		{
			name:           "qualified module member survives local shadow",
			candidate:      "Limit",
			lines:          []string{"Public Sub Run()", "  Main.Limit = 2", "End Sub"},
			procedureRange: []vbaast.Range{{StartLine: 1, StartColumn: 1, EndLine: 3, EndColumn: 1, StartByte: 41}},
			shadowed:       map[int]bool{41: true},
			moduleName:     "Main",
			wantReferenced: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			index := buildUnusedPrivateConstLexicalIndex(tt.lines, []string{tt.candidate})
			got := privateConstReferencedIndexed(tt.lines, index, tt.moduleName, tt.candidate, tt.declarator, tt.procedureRange, tt.shadowed)
			wantRegexp := legacyPrivateConstReferencedForTest(tt.lines, tt.moduleName, tt.candidate, tt.declarator, tt.procedureRange, tt.shadowed)
			if got != wantRegexp {
				t.Fatalf("indexed reference = %t, regexp reference = %t", got, wantRegexp)
			}
			if got != tt.wantReferenced {
				t.Fatalf("indexed reference = %t, want %t", got, tt.wantReferenced)
			}
		})
	}
}

func TestUnusedPrivateConstLexicalIndexScansSourceOnceForAllCandidates(t *testing.T) {
	t.Parallel()
	const lineCount = 300
	lines := make([]string, lineCount)
	names := make([]string, lineCount)
	for i := range lineCount {
		names[i] = fmt.Sprintf("Candidate%d", i)
		lines[i] = fmt.Sprintf("Private Const %s As Long = %d", names[i], i)
	}

	oneCandidate := buildUnusedPrivateConstLexicalIndex(lines, names[:1])
	allCandidates := buildUnusedPrivateConstLexicalIndex(lines, names)
	if oneCandidate.stats.sourceLinesScanned != lineCount || allCandidates.stats.sourceLinesScanned != lineCount {
		t.Fatalf("source line scans = %d and %d, want one %d-line pass each", oneCandidate.stats.sourceLinesScanned, allCandidates.stats.sourceLinesScanned, lineCount)
	}
	if oneCandidate.stats.asciiTokensScanned != allCandidates.stats.asciiTokensScanned {
		t.Fatalf("ASCII token scans changed with candidate count: one=%d all=%d", oneCandidate.stats.asciiTokensScanned, allCandidates.stats.asciiTokensScanned)
	}
	if allCandidates.stats.fallbackRegexLineScan != 0 {
		t.Fatalf("ordinary ASCII constants used regexp fallback for %d lines", allCandidates.stats.fallbackRegexLineScan)
	}
}

func TestUnusedPrivateConstLexicalIndexUsesRegexpForUnicodeCaseFolds(t *testing.T) {
	t.Parallel()
	lines := []string{"Debug.Print \"\u212a\"", "Debug.Print \"\u017f\""}
	index := buildUnusedPrivateConstLexicalIndex(lines, []string{"K", "S"})
	if got, want := index.stats.fallbackRegexLineScan, len(lines)*2; got != want {
		t.Fatalf("Unicode fold fallback scanned %d regexp lines, want %d", got, want)
	}
}

func BenchmarkUnusedPrivateConstReferenceIndex(b *testing.B) {
	const lineCount = 3000
	const candidateCount = 256
	lines := make([]string, lineCount)
	names := make([]string, candidateCount)
	for i := range names {
		names[i] = fmt.Sprintf("Candidate%d", i)
	}
	for i := range lines {
		lines[i] = fmt.Sprintf("Debug.Print Candidate%d, Candidate%d, ' Candidate%d", i%candidateCount, (i+1)%candidateCount, (i+2)%candidateCount)
	}

	b.Run("indexed", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			_ = buildUnusedPrivateConstLexicalIndex(lines, names)
		}
	})
	b.Run("regexp-per-name", func(b *testing.B) {
		matchers := make([]*regexp.Regexp, len(names))
		for i, name := range names {
			matchers[i] = unusedNameMatcher(name)
		}
		b.ReportAllocs()
		for b.Loop() {
			for _, matcher := range matchers {
				for _, line := range lines {
					_ = matcher.FindAllStringIndex(line, -1)
				}
			}
		}
	})
}

func legacyPrivateConstReferencedForTest(lines []string, moduleName, name string, declarator vbaast.Range, procedureRanges []vbaast.Range, shadowed map[int]bool) bool {
	matcher := unusedNameMatcher(name)
	for lineIndex, line := range lines {
		lineNumber := lineIndex + 1
		for _, location := range matcher.FindAllStringIndex(line, -1) {
			begin := location[0]
			if !isIdentifierByte(line[begin]) {
				begin++
			}
			if positionInRange(lineNumber, begin+1, declarator) {
				continue
			}
			shadowedHere := false
			for _, procedureRange := range procedureRanges {
				if !positionInRange(lineNumber, begin+1, procedureRange) {
					continue
				}
				if shadowed[procedureRange.StartByte] {
					shadowedHere = true
				}
				break
			}
			if shadowedHere && !qualifiedModuleReference(line[:begin], moduleName) {
				continue
			}
			return true
		}
	}
	return false
}

func TestUnusedPrivateConstFindingsCountsAuthoredLexicalMentions(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Private Const CommentMention As Long = 1
Private Const StringMention As Long = 2
Private Const InactiveMention As Long = 3
' CommentMention is mentioned in a comment.
Public Sub Run()
  Debug.Print "StringMention"
#If False Then
  Debug.Print InactiveMention
#End If
End Sub
`
	findings := runUnusedDeclarationAnalysis(t, map[string]string{"Main.bas": source})
	if got := findingsByCode(findings, "VBA266"); len(got) != 0 {
		t.Fatalf("VBA266 findings = %+v, want no finding for authored lexical mentions", got)
	}
}
