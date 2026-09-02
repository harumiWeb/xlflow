package analyze

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"sync"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var benchmarkModuleFactsOperationCount int

func TestArrayOptionPrivateModuleNormalization(t *testing.T) {
	for _, test := range []struct {
		name           string
		lines          []string
		document       procedureir.DocumentIR
		want           moduleOptionState
		wantLegacyBool bool
	}{
		{
			name:           "present",
			lines:          []string{"Option Private Module"},
			want:           moduleOptionPresent,
			wantLegacyBool: true,
		},
		{
			name:           "case and whitespace insensitive",
			lines:          []string{"\toption\tprivate\tmodule  "},
			want:           moduleOptionPresent,
			wantLegacyBool: true,
		},
		{
			name:           "trailing comment",
			lines:          []string{"Option Private Module ' module visibility"},
			want:           moduleOptionPresent,
			wantLegacyBool: true,
		},
		{
			name:           "comment only",
			lines:          []string{"' Option Private Module"},
			want:           moduleOptionAbsent,
			wantLegacyBool: false,
		},
		{
			name:           "quoted text",
			lines:          []string{"Debug.Print \"Option Private Module\""},
			want:           moduleOptionAbsent,
			wantLegacyBool: false,
		},
		{
			name:           "incomplete option",
			lines:          []string{"Option Private"},
			want:           moduleOptionUnknown,
			wantLegacyBool: false,
		},
		{
			name:           "recovered declaration elsewhere",
			lines:          []string{"Option Private Module"},
			document:       procedureir.DocumentIR{Parse: procedureir.ParseSummary{HasMissing: true}},
			want:           moduleOptionUnknown,
			wantLegacyBool: true,
		},
		{
			name:           "syntax error elsewhere",
			lines:          []string{"Option Private Module"},
			document:       procedureir.DocumentIR{Parse: procedureir.ParseSummary{HasError: true}},
			want:           moduleOptionUnknown,
			wantLegacyBool: true,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			facts := buildModuleAnalysisFacts(test.lines, test.document, nil)
			if got := facts.privateModuleState(); got != test.want {
				t.Fatalf("privateModuleState(%q) = %v, want %v", test.lines, got, test.want)
			}
			if got := arrayOptionPrivateModule(test.lines); got != test.wantLegacyBool {
				t.Fatalf("arrayOptionPrivateModule(%q) = %v, want %v", test.lines, got, test.wantLegacyBool)
			}
		})
	}
}

func TestModuleAnalysisFactsConcurrentRead(t *testing.T) {
	lines := []string{
		"Option Private Module",
		"Private Const ModuleValue As Long = 1",
		"moduleArray = Array(1)",
		"Public Sub First()",
		"    Dim localValue As Long",
		"End Sub",
		"Public Sub Second()",
		"    Dim otherValue As Long",
		"End Sub",
	}
	procedures := []sourceProcedure{
		{Name: "First", StartLine: 4, EndLine: 6, StartByte: 101},
		{Name: "Second", StartLine: 7, EndLine: 9, StartByte: 201},
	}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, procedures)
	if facts == nil {
		t.Fatal("buildModuleAnalysisFacts returned nil")
	}

	const readers = 32
	const iterations = 100
	var wg sync.WaitGroup
	wg.Add(readers)
	for reader := 0; reader < readers; reader++ {
		go func() {
			defer wg.Done()
			for iteration := 0; iteration < iterations; iteration++ {
				if facts.privateModuleState() != moduleOptionPresent || !facts.privateModulePresent() {
					t.Errorf("private module state is not stable")
					return
				}
				if !facts.lineInProcedure(4) || facts.lineInProcedure(1) {
					t.Errorf("unexpected procedure ownership")
					return
				}
				declarations := facts.procedureDeclarations(procedures[0])
				if got := declarations["localvalue"].Name; got != "localValue" {
					t.Errorf("procedure declarations = %#v", declarations)
					return
				}
				if facts.procedureFactsFor(procedures[0]) == nil || facts.procedureFactsFor(procedures[1]) == nil {
					t.Errorf("procedure facts lookup returned nil")
					return
				}
				if !facts.hasConstant("modulevalue") || facts.hasConstantForProcedure("LocalValue", procedures[0]) || facts.hasConstantForProcedure("OtherValue", procedures[0]) {
					t.Errorf("constant lookup returned inconsistent result")
					return
				}
				if !facts.hasProcedure("FIRST") || !facts.hasProcedure("second") {
					t.Errorf("procedure lookup returned inconsistent result")
					return
				}
				operationCount := 0
				facts.forEachArrayOperationFor("MODULEARRAY", func(operation moduleArrayOperationFact) {
					if operation.Kind == moduleArrayWholeAssignment && operation.RHS == "Array(1)" {
						operationCount++
					}
				})
				if operationCount != 1 {
					t.Errorf("module operation lookup count = %d, want one", operationCount)
					return
				}
			}
		}()
	}
	wg.Wait()
}

func TestModuleAnalysisFactsIndexesArrayOperations(t *testing.T) {
	lines := []string{
		"Private values() As Long",
		"Private other() As Long",
		"values = Array(1)",
		"ReDim values(0 To 2)",
		"ReDim Preserve values(0 To 4)",
		"Erase values, other",
	}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, nil)

	var values []moduleArrayOperationFact
	facts.forEachArrayOperationFor("VALUES", func(operation moduleArrayOperationFact) {
		values = append(values, operation)
	})
	if len(values) != 4 {
		t.Fatalf("values operations = %#v, want whole assignment, two ReDim operations, and Erase", values)
	}
	if values[0].Kind != moduleArrayWholeAssignment || values[0].RHS != "Array(1)" {
		t.Fatalf("whole-array operation = %#v", values[0])
	}
	if values[1].Kind != moduleArrayDirectRedim || values[1].Preserve || values[1].Dimensions != "0 To 2" {
		t.Fatalf("direct ReDim operation = %#v", values[1])
	}
	if values[2].Kind != moduleArrayDirectRedim || !values[2].Preserve || values[2].Dimensions != "0 To 4" {
		t.Fatalf("Preserve ReDim operation = %#v", values[2])
	}
	if values[3].Kind != moduleArrayErase {
		t.Fatalf("Erase operation = %#v", values[3])
	}

	var other []moduleArrayOperationFact
	facts.forEachArrayOperationFor("other", func(operation moduleArrayOperationFact) {
		other = append(other, operation)
	})
	if len(other) != 1 || other[0].Kind != moduleArrayErase {
		t.Fatalf("other operations = %#v, want one Erase", other)
	}

	var atLine []moduleArrayOperationFact
	facts.forEachArrayOperationAt(values[1].Line, func(operation moduleArrayOperationFact) {
		atLine = append(atLine, operation)
	})
	if len(atLine) != 1 || atLine[0].Name != "values" {
		t.Fatalf("line operation index = %#v, want one values operation", atLine)
	}
}

func TestModuleArrayOperationIndexMatchesLegacySourceScan(t *testing.T) {
	lines := []string{
		"ReDim first(0 To 1), second(0 To 2) ' multi-clause",
		"ReDim Preserve first(0 To 3), second(0 To 4)",
		"Erase first, second, indexed(1)",
		"first(0) = 1",
		"first = Array(1) ' whole-array assignment",
		"first% = Array(2) ' type-suffix identifier",
	}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, nil)
	got := make([]moduleArrayOperationFact, 0)
	for _, name := range []string{"first", "second", "indexed"} {
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			got = append(got, operation)
		})
	}
	want := legacyModuleArrayOperationFacts(lines)
	canonicalize := func(operations []moduleArrayOperationFact) {
		sort.SliceStable(operations, func(i, j int) bool {
			if operations[i].Line != operations[j].Line {
				return operations[i].Line < operations[j].Line
			}
			if operations[i].Kind != operations[j].Kind {
				return operations[i].Kind < operations[j].Kind
			}
			return operations[i].Name < operations[j].Name
		})
	}
	canonicalize(got)
	canonicalize(want)
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("module operation index differs from legacy scan:\n got: %#v\nwant: %#v", got, want)
	}
}

func legacyModuleArrayOperationFacts(lines []string) []moduleArrayOperationFact {
	operations := make([]moduleArrayOperationFact, 0)
	for lineNo, line := range lines {
		text := normalizedCodeLine(line)
		if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed {
			operations = append(operations, moduleArrayOperationFact{
				Name: strings.ToLower(cleanIdentifier(lhs)), Line: lineNo,
				RHS: strings.TrimSpace(rhs), Kind: moduleArrayWholeAssignment,
			})
		}
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
			preserve := strings.TrimSpace(match[1]) != ""
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct {
					continue
				}
				operations = append(operations, moduleArrayOperationFact{
					Name: strings.ToLower(cleanIdentifier(redim.name)), Line: lineNo,
					Dimensions: redim.dimensions, Preserve: preserve,
					Kind: moduleArrayDirectRedim,
				})
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
			for _, target := range splitArgs(match[1]) {
				name := strings.ToLower(strings.TrimSpace(target))
				if !arrayEraseNameRe.MatchString(name) {
					continue
				}
				operations = append(operations, moduleArrayOperationFact{Name: name, Line: lineNo, Kind: moduleArrayErase})
			}
		}
	}
	return operations
}

func TestModuleIdempotentSetupPreservesNonCandidateExecutableTail(t *testing.T) {
	for _, test := range []struct {
		name       string
		tail       string
		wantValues bool
	}{
		{name: "assignment is final", tail: "", wantValues: true},
		{name: "debug statement follows assignment", tail: "Debug.Print \"initialized\"", wantValues: false},
	} {
		t.Run(test.name, func(t *testing.T) {
			lines := []string{
				"Private ready As Boolean",
				"Private values() As Long",
				"Private Sub EnsureValues()",
				"    If ready Then Exit Sub",
				"    ReDim values(0 To 1)",
				"    ready = True",
			}
			if test.tail != "" {
				lines = append(lines, "    "+test.tail)
			}
			lines = append(lines, "End Sub")
			proc := sourceProcedure{Name: "EnsureValues", StartLine: 3, EndLine: len(lines), StartByte: 1}
			facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, []sourceProcedure{proc})
			file := parsedFile{Lines: lines, ModuleFacts: facts}
			moduleDecls := map[string]sourceDeclaration{
				"ready":  {Name: "ready", Type: "Boolean"},
				"values": {Name: "values", Array: true, Type: "Long"},
			}
			got := arrayModuleIdempotentSetupArrays(file, proc, moduleDecls, analysisContext{})
			if got["values"] != test.wantValues {
				t.Fatalf("idempotent setup summary = %#v, values presence = %v, want %v", got, got["values"], test.wantValues)
			}
		})
	}
}

func TestModuleAnalysisFactsSmallAllocationGuard(t *testing.T) {
	if !moduleFactsAllocationGuardEnabled {
		t.Skip("allocation guard is disabled under the race detector")
	}
	lines := make([]string, 8*3)
	procedures := make([]sourceProcedure, 0, 8)
	for index := 0; index < 8; index++ {
		start := index*3 + 1
		lines[start-1] = "Private Sub Run" + strconvItoa(index) + "()"
		lines[start] = "    Dim localValue As Long"
		lines[start+1] = "End Sub"
		procedures = append(procedures, sourceProcedure{StartLine: start, EndLine: start + 2, StartByte: index + 1})
	}
	declarations := make([]procedureir.Declaration, 0, 4)
	for index := 0; index < 4; index++ {
		declarations = append(declarations, procedureir.Declaration{
			Name: "ModuleValue" + strconvItoa(index), Type: "Long", Scope: procedureir.ScopeModule,
			Kind: "variable", Range: vbaast.Range{StartLine: index + 1, EndLine: index + 1},
		})
	}
	document := procedureir.DocumentIR{Declarations: declarations}
	result := testing.Benchmark(func(b *testing.B) {
		for index := 0; index < b.N; index++ {
			benchmarkModuleFactsSink = buildModuleAnalysisFacts(lines, document, procedures)
		}
	})
	// Baseline is the pre-#711 small workload measured on the repository's
	// Windows Go toolchain. Keep the guard at +2% to catch accidental per-file
	// normalization regressions while allowing normal benchmark jitter.
	const baselineBytes = 60560
	const baselineAllocs = 295
	if got := result.AllocedBytesPerOp(); got > baselineBytes*102/100 {
		t.Fatalf("small module allocation = %d B/op, baseline %d B/op (+2%% guard)", got, baselineBytes)
	}
	if got := result.AllocsPerOp(); got > baselineAllocs*102/100 {
		t.Fatalf("small module allocations = %d allocs/op, baseline %d (+2%% guard)", got, baselineAllocs)
	}
}

func TestModuleOptionScanCountDoesNotScaleWithPublicProcedures(t *testing.T) {
	for _, procedureCount := range []int{1, 100, 1000} {
		t.Run(fmt.Sprintf("procedures_%d", procedureCount), func(t *testing.T) {
			values := runModuleOptionScanBenchmarkAnalysis(t, procedureCount)
			if got := values[analysisstats.ModuleOptionScansCounter]; got != 1 {
				t.Fatalf("module option scans = %d for %d procedures, want one", got, procedureCount)
			}
			if got := values[analysisstats.ModuleFactBuildsCounter]; got != 1 {
				t.Fatalf("module fact builds = %d for %d procedures, want one", got, procedureCount)
			}
		})
	}
}

func runModuleOptionScanBenchmarkAnalysis(t *testing.T, procedureCount int) map[string]uint64 {
	t.Helper()
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Option Explicit\nOption Private Module\n"
	for index := 0; index < procedureCount; index++ {
		source += fmt.Sprintf("Public Sub Run%d()\nEnd Sub\n", index)
	}
	if err := os.WriteFile(filepath.Join(modules, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	recorder := analysisstats.NewRecorder()
	_, err := (Analyzer{RootDir: root, Config: config.Default()}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder))
	if err != nil {
		t.Fatal(err)
	}
	_, counters := recorder.Totals()
	values := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		values[counter.Name] = counter.Value
	}
	return values
}

// BenchmarkModuleAnalysisFactsOperationHeavy exercises the profiler-driven
// operation index with the same repeated setup idiom that procedure workers
// query. The benchmark intentionally measures construction only; accessors
// are allocation-free reads over the frozen facts object.
func BenchmarkModuleAnalysisFactsOperationHeavy(b *testing.B) {
	const procedureCount = 500
	lines := []string{
		"Option Private Module",
		"Private ready As Boolean",
		"Private values() As Long",
		"Private other() As Long",
	}
	procedures := make([]sourceProcedure, 0, procedureCount)
	for index := 0; index < procedureCount; index++ {
		start := len(lines) + 1
		lines = append(lines,
			fmt.Sprintf("Public Sub Run%d()", index),
			"    If ready Then Exit Sub",
			fmt.Sprintf("    ReDim values(0 To %d)", index%8+1),
			"    ready = True",
			"    Erase values, other",
			"End Sub",
		)
		procedures = append(procedures, sourceProcedure{
			StartLine: start, EndLine: start + 5, StartByte: index + 1,
		})
	}
	document := procedureir.DocumentIR{}
	b.ReportAllocs()
	b.ResetTimer()
	for index := 0; index < b.N; index++ {
		benchmarkModuleFactsSink = buildModuleAnalysisFacts(lines, document, procedures)
	}
	b.ReportMetric(float64(procedureCount), "procedures/op")
}

func BenchmarkModuleAnalysisFactsOperationAccessors(b *testing.B) {
	lines := []string{
		"Private ready As Boolean",
		"Private values() As Long",
		"If ready Then Exit Sub",
		"ReDim values(0 To 1)",
		"ready = True",
		"Erase values",
	}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, nil)
	b.ReportAllocs()
	b.ResetTimer()
	for index := 0; index < b.N; index++ {
		count := 0
		facts.forEachArrayOperationFor("values", func(moduleArrayOperationFact) { count++ })
		facts.forEachArrayOperationAt(3, func(moduleArrayOperationFact) { count++ })
		benchmarkModuleFactsOperationCount += count
		benchmarkModuleFactsSink = facts
	}
}
