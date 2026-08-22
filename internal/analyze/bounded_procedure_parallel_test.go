package analyze

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"path/filepath"
	"reflect"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
)

// TestAnalyzerSingleModuleProcedureAnalysisIsDeterministic exercises the
// workload shape that file-level workers cannot parallelize: one source file
// with hundreds of independent procedures. The serial and bounded paths must
// produce identical findings, including their serialized representation.
func TestAnalyzerSingleModuleProcedureAnalysisIsDeterministic(t *testing.T) {
	root := t.TempDir()
	fixture := writeSingleModuleBenchmarkProject(t, root, singleModuleBenchmarkWorkload{shape: "independent", size: 500})
	if fixture.procedures != 500 {
		t.Fatalf("single-module fixture has %d procedures, want 500", fixture.procedures)
	}
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	cfg := config.Default()

	serial, err := (Analyzer{RootDir: root, Config: cfg, analysisWorkerLimit: 1}).RunResult()
	if err != nil {
		t.Fatalf("serial analysis: %v", err)
	}
	parallelAnalyzer := Analyzer{RootDir: root, Config: cfg, analysisWorkerLimit: 4}
	parallel, err := parallelAnalyzer.RunResult()
	if err != nil {
		t.Fatalf("bounded analysis: %v", err)
	}
	repeated, err := parallelAnalyzer.RunResult()
	if err != nil {
		t.Fatalf("repeated bounded analysis: %v", err)
	}

	if !reflect.DeepEqual(serial, parallel) {
		t.Fatalf("serial and bounded results differ:\nserial:   %+v\nbounded: %+v", serial, parallel)
	}
	if !reflect.DeepEqual(parallel, repeated) {
		t.Fatalf("repeated bounded results differ:\nfirst:  %+v\nsecond: %+v", parallel, repeated)
	}
	serialJSON, err := json.Marshal(serial)
	if err != nil {
		t.Fatalf("marshal serial result: %v", err)
	}
	parallelJSON, err := json.Marshal(parallel)
	if err != nil {
		t.Fatalf("marshal bounded result: %v", err)
	}
	if !bytes.Equal(serialJSON, parallelJSON) {
		t.Fatalf("serial and bounded JSON differ:\nserial:   %s\nbounded: %s", serialJSON, parallelJSON)
	}
}

// TestAnalyzeFilesBoundedCapsConcurrentAnalyses protects the existing
// process-wide file budget. Procedure-level scheduling is allowed to add work
// inside a file, but it must not multiply the outer worker limit.
func TestAnalyzeFilesBoundedCapsConcurrentAnalyses(t *testing.T) {
	for _, workerLimit := range []int{1, 2, 4} {
		workerLimit := workerLimit
		t.Run(fmt.Sprintf("limit-%d", workerLimit), func(t *testing.T) {
			files := make([]parsedFile, 8)
			for index := range files {
				files[index].Path = fmt.Sprintf("Module%d.bas", index)
			}

			var active atomic.Int32
			var maxActive atomic.Int32
			updateMax := func(value int32) {
				for {
					old := maxActive.Load()
					if value <= old || maxActive.CompareAndSwap(old, value) {
						return
					}
				}
			}
			firstBatchStarted := make(chan struct{})
			release := make(chan struct{})
			var firstBatchOnce sync.Once
			resultCh := make(chan error, 1)
			go func() {
				_, err := (Analyzer{analysisWorkerLimit: workerLimit}).analyzeFilesBoundedWith(
					context.Background(), files, analysisContext{}, effects.ProjectSummary{}, nil,
					func(ctx context.Context, file parsedFile, _ analysisContext, _ effects.ProjectSummary, _ *apiTypeIndex) ([]Finding, []Finding, error) {
						current := active.Add(1)
						defer active.Add(-1)
						updateMax(current)
						if current == int32(workerLimit) {
							firstBatchOnce.Do(func() { close(firstBatchStarted) })
						}
						select {
						case <-release:
						case <-ctx.Done():
							return nil, nil, ctx.Err()
						}
						return nil, nil, nil
					},
				)
				resultCh <- err
			}()

			select {
			case <-firstBatchStarted:
			case <-time.After(5 * time.Second):
				t.Fatal("timed out waiting for the bounded worker batch")
			}
			close(release)
			select {
			case err := <-resultCh:
				if err != nil {
					t.Fatalf("bounded analysis: %v", err)
				}
			case <-time.After(5 * time.Second):
				t.Fatal("timed out waiting for bounded analysis completion")
			}
			if got := maxActive.Load(); got > int32(workerLimit) {
				t.Fatalf("maximum concurrent analyses = %d, want <= %d", got, workerLimit)
			}
		})
	}
}

func TestAnalyzerSingleModuleProcedureCancellationReturnsNoResult(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()

	result, err := (Analyzer{RootDir: t.TempDir(), Config: config.Default(), analysisWorkerLimit: 4}).RunResultContext(ctx)
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("RunResultContext error = %v, want context.Canceled", err)
	}
	if !reflect.DeepEqual(result, Result{}) {
		t.Fatalf("canceled analysis returned partial result: %+v", result)
	}
}

func TestAnalyzeFilesBoundedSkipsProcedurePoolForSmallFiles(t *testing.T) {
	files := []parsedFile{{Procedures: make([]sourceProcedure, procedureParallelThreshold-1)}}
	seen := make(chan bool, 1)
	_, err := (Analyzer{analysisWorkerLimit: 4}).analyzeFilesBoundedWith(
		context.Background(), files, analysisContext{}, effects.ProjectSummary{}, nil,
		func(ctx context.Context, _ parsedFile, _ analysisContext, _ effects.ProjectSummary, _ *apiTypeIndex) ([]Finding, []Finding, error) {
			budget := analysisExecutionBudgetFromContext(ctx)
			seen <- budget != nil && budget.procedureJobs != nil
			return nil, nil, nil
		},
	)
	if err != nil {
		t.Fatalf("bounded analysis: %v", err)
	}
	if got := <-seen; got {
		t.Fatal("small-file analysis started a procedure pool")
	}
}
