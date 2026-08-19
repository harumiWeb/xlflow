package analyze

import (
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestAnalyzerBoundedFileAnalysisIsDeterministic(t *testing.T) {
	root := t.TempDir()
	for index, name := range []string{"First.bas", "Second.bas", "Third.bas", "Fourth.bas"} {
		writeModule(t, root, name, "Option Explicit\n\nPublic Sub Probe"+string(rune('A'+index))+"()\n    Debug.Print \"probe\"\nEnd Sub\n")
	}

	cfg := config.Default()
	serial := Analyzer{RootDir: root, Config: cfg, analysisWorkerLimit: 1}
	parallel := Analyzer{RootDir: root, Config: cfg, analysisWorkerLimit: 4}
	serialResult, err := serial.RunResult()
	if err != nil {
		t.Fatalf("serial analysis: %v", err)
	}
	parallelResult, err := parallel.RunResult()
	if err != nil {
		t.Fatalf("parallel analysis: %v", err)
	}
	if !reflect.DeepEqual(serialResult, parallelResult) {
		t.Fatalf("serial and parallel results differ:\nserial:   %+v\nparallel: %+v", serialResult, parallelResult)
	}
}
