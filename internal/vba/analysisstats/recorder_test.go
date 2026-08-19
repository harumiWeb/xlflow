package analysisstats

import (
	"context"
	"errors"
	"reflect"
	"testing"
	"time"
)

func TestMeasureWaitPopulatesWaitDuration(t *testing.T) {
	recorder := NewRecorder()
	ctx := WithRecorder(context.Background(), recorder)
	finish := MeasureWait(ctx, "procedure_ir_singleflight")
	time.Sleep(time.Millisecond)
	finish(nil)

	stages, _ := recorder.Snapshot()
	if len(stages) != 1 || stages[0].Name != "procedure_ir_singleflight" || stages[0].Wait <= 0 || stages[0].Elapsed != 0 || stages[0].Outcome != "ok" {
		t.Fatalf("wait stage = %+v", stages)
	}
}

func TestMeasureWithoutRecorderIsNoOp(t *testing.T) {
	finish := Measure(context.Background(), "parse")
	finish(1, nil)
	finishWait := MeasureWait(context.Background(), "procedure_ir_singleflight")
	finishWait(nil)
}

func TestTotalsAggregateStagesAndSortCounters(t *testing.T) {
	recorder := NewRecorder()
	recorder.Record(Stage{Name: "parse", Elapsed: time.Millisecond, Outcome: "ok", ResultCount: 1})
	recorder.Record(Stage{Name: "cfg", Elapsed: 2 * time.Millisecond, Outcome: "ok", ResultCount: 2})
	recorder.Record(Stage{Name: "parse", Elapsed: 3 * time.Millisecond, Outcome: "error", ResultCount: 4})
	recorder.Add("procedure_count", 3)
	recorder.AddSum("file_count", 1)
	recorder.AddSum("procedure_count", 2)
	recorder.AddMax("max_lines_per_file", 11)
	recorder.AddMax("max_lines_per_file", 7)
	recorder.AddMax("max_procedures_per_file", 2)
	recorder.AddMax("max_procedures_per_file", 5)

	stages, counters := recorder.Totals()
	if len(stages) != 2 {
		t.Fatalf("stage totals = %+v", stages)
	}
	if got := stages[0]; got.Name != "parse" || got.Elapsed != 4*time.Millisecond || got.ResultCount != 5 || got.Calls != 2 || got.Outcome != "error" {
		t.Fatalf("parse total = %+v", got)
	}
	if got := stages[1]; got.Name != "cfg" || got.Calls != 1 || got.Outcome != "ok" {
		t.Fatalf("cfg total = %+v", got)
	}
	wantCounters := []Counter{
		{Name: "file_count", Value: 1},
		{Name: "max_lines_per_file", Value: 11},
		{Name: "max_procedures_per_file", Value: 5},
		{Name: "procedure_count", Value: 5},
	}
	if !reflect.DeepEqual(counters, wantCounters) {
		t.Fatalf("counters = %+v, want %+v", counters, wantCounters)
	}
}

func TestMeasureRecordsCanceledOutcome(t *testing.T) {
	recorder := NewRecorder()
	ctx, cancel := context.WithCancel(WithRecorder(context.Background(), recorder))
	finish := Measure(ctx, "parse")
	cancel()
	finish(0, errors.New("stopped"))

	stages, _ := recorder.Totals()
	if len(stages) != 1 || stages[0].Outcome != "canceled" {
		t.Fatalf("stages = %+v", stages)
	}
}
