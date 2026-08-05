package analysisstats

import (
	"context"
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
