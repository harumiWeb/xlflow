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
	b.Run("many-procedure-prefix-lookups", func(b *testing.B) {
		index := buildLoopInvariantStringConstantIndex(lines)
		b.ReportAllocs()
		for b.Loop() {
			for line := 1; line <= lineCount; line++ {
				constants := loopInvariantStringConstantSet{index: index, prefixLine: line}
				_, _ = constants.lookup(names[line%constantCount])
			}
		}
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
