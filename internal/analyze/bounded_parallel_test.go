package analyze

import (
	"context"
	"errors"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
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

func TestAnalyzeFilesBoundedPreservesWorkerError(t *testing.T) {
	lowStarted := make(chan struct{})
	highStarted := make(chan struct{})
	releaseLow := make(chan struct{})
	workerFailure := errors.New("worker failure")

	resultCh := make(chan error, 1)
	go func() {
		_, err := (Analyzer{analysisWorkerLimit: 2}).analyzeFilesBoundedWith(
			context.Background(),
			[]parsedFile{{Path: "low"}, {Path: "high"}},
			analysisContext{}, effects.ProjectSummary{}, nil,
			func(ctx context.Context, file parsedFile, _ analysisContext, _ effects.ProjectSummary, _ *apiTypeIndex) ([]Finding, []Finding, error) {
				switch file.Path {
				case "low":
					close(lowStarted)
					<-releaseLow
					return nil, nil, ctx.Err()
				case "high":
					close(highStarted)
					return nil, nil, workerFailure
				default:
					return nil, nil, nil
				}
			},
		)
		resultCh <- err
	}()

	<-lowStarted
	<-highStarted
	close(releaseLow)
	if err := <-resultCh; !errors.Is(err, workerFailure) {
		t.Fatalf("bounded analysis error = %v, want originating worker error %v", err, workerFailure)
	}
}
