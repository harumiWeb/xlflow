package analyze

import (
	"context"
	"errors"
	"testing"
)

func TestProcedureSemanticResultStoreDoesNotCacheBuildError(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	store := procedureSemanticResultStore{}

	_, err := store.materializeArray(ctx, Analyzer{}, parsedFile{}, sourceProcedure{}, analysisContext{}, nil, procedureAnalysisPlan{})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("first materialization error = %v, want context.Canceled", err)
	}
	if store.arrayBuilt {
		t.Fatal("array result store marked a failed build as complete")
	}

	_, err = store.materializeArray(ctx, Analyzer{}, parsedFile{}, sourceProcedure{}, analysisContext{}, nil, procedureAnalysisPlan{})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("second materialization error = %v, want context.Canceled", err)
	}
}
