package analyze

import (
	"fmt"
	"strings"
	"sync"
	"testing"

	"github.com/harumiWeb/xlflow/internal/gui"
)

func TestLoopInvariantStringConstantIndexConcurrentInitialization(t *testing.T) {
	facts := &moduleAnalysisFacts{}
	lines := []string{`Private Const Selector As String = "Worksheets"`}
	results := make([]*loopInvariantStringConstantIndex, 16)
	var workers sync.WaitGroup
	for i := range results {
		workers.Go(func() { results[i] = facts.loopInvariantStringConstantIndexFor(lines) })
	}
	workers.Wait()
	for _, result := range results {
		if result != results[0] || result.linesScanned != len(lines) {
			t.Fatal("concurrent readers did not share one immutable module index")
		}
	}
}

func TestLoopInvariantStringConstantIndexCountsVisitedLines(t *testing.T) {
	for _, size := range []int{0, 1, 10000} {
		t.Run(fmt.Sprintf("lines-%d", size), func(t *testing.T) {
			// Nonmatching lines are still visited by the index build.
			index := buildLoopInvariantStringConstantIndex(make([]string, size))
			if index.linesScanned != size {
				t.Fatalf("visited %d lines, want %d", index.linesScanned, size)
			}
		})
	}
}

func TestLoopInvariantStringConstantIndexReusesModuleScan(t *testing.T) {
	t.Parallel()
	lines := []string{
		`Public Const Selector As String = "Worksheets"`,
		"Public Sub First()",
		`  Const Selector As String = "Names"`,
		"End Sub",
		"Public Sub Second()",
		`  Const FutureSelector As String = "Charts"`,
		"End Sub",
	}
	facts := &moduleAnalysisFacts{}
	file := parsedFile{Lines: lines, ModuleFacts: facts}
	procedures := []sourceProcedure{{StartLine: 2}, {StartLine: 5}}

	for _, proc := range procedures {
		got := loopInvariantStringConstants(file, proc)
		want := legacyLoopInvariantStringConstantsForTest(file, proc)
		for _, name := range []string{"selector", "futureselector"} {
			gotValue, gotOK := got.lookup(name)
			wantValue, wantOK := want[name]
			if gotValue != wantValue || gotOK != wantOK {
				t.Fatalf("prefix %d lookup %q = %q, %t; legacy = %q, %t", proc.StartLine, name, gotValue, gotOK, wantValue, wantOK)
			}
		}
	}

	first := loopInvariantStringConstants(file, procedures[0])
	second := loopInvariantStringConstants(file, procedures[1])
	if first.index != second.index {
		t.Fatal("procedure lanes did not reuse the module/revision constant index")
	}
	if got := facts.loopInvariantStringConstIndex.linesScanned; got != len(lines) {
		t.Fatalf("module source lines scanned = %d, want one %d-line index build", got, len(lines))
	}
}

func TestLoopInvariantStringConstantSetPreservesPrefixOrderAndLocalOverride(t *testing.T) {
	t.Parallel()
	index := &loopInvariantStringConstantIndex{byName: map[string][]loopInvariantStringConstantFact{
		"selector": {
			{line: 1, value: `"Worksheets"`},
			{line: 3, value: `"Names"`},
		},
		"futureselector": {{line: 6, value: `"Charts"`}},
	}}
	prefix := loopInvariantStringConstantSet{index: index, prefixLine: 5}
	if got, ok := prefix.lookup("selector"); !ok || got != `"Names"` {
		t.Fatalf("later procedure prefix selector = %q, %t; want latest source-order value", got, ok)
	}
	if _, ok := prefix.lookup("futureselector"); ok {
		t.Fatal("a declaration after the procedure prefix was visible")
	}
	withLocal := loopInvariantStringConstantSet{
		index:      index,
		prefixLine: 5,
		procedure:  map[string]string{"selector": `"Sheets"`},
	}
	if got, ok := withLocal.lookup("selector"); !ok || got != `"Sheets"` {
		t.Fatalf("procedure-local override = %q, %t; want local value", got, ok)
	}
}

func BenchmarkLoopInvariantStringConstantIndex(b *testing.B) {
	const lineCount = 3000
	const constantCount = 256
	lines := make([]string, lineCount)
	names := make([]string, constantCount)
	for i := range names {
		names[i] = fmt.Sprintf("selector%d", i)
	}
	for i := range lines {
		if i < constantCount {
			lines[i] = fmt.Sprintf(`Private Const Selector%d As String = "Worksheets"`, i)
			continue
		}
		lines[i] = fmt.Sprintf("Debug.Print value%d", i)
	}

	b.Run("index-build", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			_ = buildLoopInvariantStringConstantIndex(lines)
		}
	})
	b.Run("existing-extraction-empty-prefix-lookups", func(b *testing.B) {
		index := buildLoopInvariantStringConstantIndex(lines)
		if len(index.byName) != 0 {
			b.Fatal("existing extraction semantics changed; update this control explicitly")
		}
		b.ReportAllocs()
		for b.Loop() {
			for line := 1; line <= lineCount; line++ {
				constants := loopInvariantStringConstantSet{index: index, prefixLine: line}
				_, _ = constants.lookup(names[line%constantCount])
			}
		}
	})
	b.Run("populated-source-order-prefix-lookups", func(b *testing.B) {
		// The production extraction guard currently produces no facts. Seed
		// immutable facts explicitly to measure the positive lookup kernel
		// without extending the rule's supported selector proofs.
		index := &loopInvariantStringConstantIndex{byName: make(map[string][]loopInvariantStringConstantFact, constantCount)}
		for _, name := range names {
			index.byName[name] = []loopInvariantStringConstantFact{
				{line: 1, value: `"Worksheets"`},
				{line: lineCount / 2, value: `"Names"`},
				{line: lineCount, value: `"Sheets"`},
			}
		}
		for _, check := range []struct {
			line int
			want string
		}{{0, ""}, {1, `"Worksheets"`}, {lineCount / 2, `"Names"`}, {lineCount, `"Sheets"`}} {
			constants := loopInvariantStringConstantSet{index: index, prefixLine: check.line}
			value, ok := constants.lookup(names[0])
			if value != check.want || ok != (check.line > 0) {
				b.Fatalf("prefix %d = %q, %t; want %q", check.line, value, ok, check.want)
			}
		}
		b.ReportAllocs()
		for b.Loop() {
			for line := 1; line <= lineCount; line++ {
				constants := loopInvariantStringConstantSet{index: index, prefixLine: line}
				if _, ok := constants.lookup(names[line%constantCount]); !ok {
					b.Fatal("populated prefix lookup missed an available fact")
				}
			}
		}
		b.ReportMetric(lineCount, "successful-lookups/op")
	})
}

func legacyLoopInvariantStringConstantsForTest(file parsedFile, proc sourceProcedure) map[string]string {
	constants := map[string]string{}
	for lineNumber, line := range file.Lines {
		if proc.StartLine > 0 && lineNumber+1 > proc.StartLine {
			break
		}
		line = strings.TrimSpace(gui.StripComment(line))
		match := loopInvariantConstStringRe.FindStringSubmatch(line)
		if len(match) == 2 {
			constants[strings.ToLower(match[1])] = match[2]
		}
	}
	for statement := range proc.Statements.All() {
		line := strings.TrimSpace(gui.StripComment(statement.Text))
		match := loopInvariantConstStringRe.FindStringSubmatch(line)
		if len(match) == 2 {
			constants[strings.ToLower(match[1])] = match[2]
		}
	}
	return constants
}
