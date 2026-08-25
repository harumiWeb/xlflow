package analyze

import (
	"context"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

func TestSemanticQueryReusesProcedureKernelsAcrossBatchRevisions(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Const ModuleConstant As Long = 1
Public Sub Run()
    Dim values() As Long
    ReDim values(1 To 2)
    values(1) = 1
End Sub

Public Sub Untouched()
    Dim other() As Long
    ReDim other(1 To 2)
    other(1) = 1
End Sub
`)
	store := semanticquery.New(semanticquery.Options{})
	run := func() (Result, map[string]uint64, error) {
		recorder := analysisstats.NewRecorder()
		ctx := analysisstats.WithRecorder(context.Background(), recorder)
		ctx = semanticquery.WithContext(ctx, semanticquery.Context{Store: store})
		result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResultContext(ctx)
		_, counters := recorder.Snapshot()
		return result, counters, err
	}
	first, firstCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	second, secondCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first.Findings, second.Findings) {
		t.Fatalf("warm analysis changed findings: first=%#v second=%#v", first.Findings, second.Findings)
	}
	if secondCounters[analysisstats.SemanticQueryHitsCounter] == 0 {
		t.Fatalf("warm analysis recorded no semantic query hits: first=%v second=%v", firstCounters, secondCounters)
	}
	if secondCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] >= firstCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] {
		t.Fatalf("warm analysis recomputed at least as many kernels: first=%v second=%v", firstCounters, secondCounters)
	}
	sourcePath := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(sourcePath)
	if err != nil {
		t.Fatal(err)
	}
	moduleUpdated := strings.Replace(string(source), "ModuleConstant As Long = 1", "ModuleConstant As Long = 2", 1)
	if moduleUpdated == string(source) {
		t.Fatal("module constant replacement did not change source")
	}
	if err := os.WriteFile(sourcePath, []byte(moduleUpdated), 0o644); err != nil {
		t.Fatal(err)
	}
	_, moduleCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	if moduleCounters[analysisstats.SemanticQueryMissesCounter] == 0 {
		t.Fatalf("module input edit reused all old kernels: counters=%v", moduleCounters)
	}
	source, err = os.ReadFile(sourcePath)
	if err != nil {
		t.Fatal(err)
	}
	updated := strings.Replace(string(source), "other(1) = 1", "other(1) = 2", 1)
	if err := os.WriteFile(sourcePath, []byte(updated), 0o644); err != nil {
		t.Fatal(err)
	}
	_, changedCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	if changedCounters[analysisstats.SemanticQueryHitsCounter] == 0 {
		t.Fatalf("local edit did not reuse the unchanged procedure: counters=%v", changedCounters)
	}
	if changedCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] >= firstCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] {
		t.Fatalf("local edit recomputed every kernel: first=%v changed=%v", firstCounters, changedCounters)
	}
}
