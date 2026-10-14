package spec

import (
	"math"
	"testing"
)

func TestInspectSnapshotRejectsInvalidClientDimensions(t *testing.T) {
	for _, field := range []string{"client_width", "client_height"} {
		for _, value := range []float64{-1, math.NaN(), math.Inf(1), float64(math.MaxInt32+1) * 72 / 2540} {
			_, err := FormSpecFromInspectSnapshot(map[string]any{
				"name": "TestForm", "basis": "designer", field: value, "controls": []any{},
			})
			if err == nil {
				t.Fatalf("snapshot accepted %s=%v", field, value)
			}
		}
	}
}

func TestInspectSnapshotSeedsClientDimensionsWithoutOuterConversion(t *testing.T) {
	snapshot, err := FormSpecFromInspectSnapshot(map[string]any{
		"name": "TestForm", "basis": "designer", "caption": "Test",
		"width": 530.0, "height": 450.0,
		"client_width": 520.0, "client_height": 420.0,
		"controls": []any{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if snapshot.Form.Width == nil || *snapshot.Form.Width != 520 || snapshot.Form.Height == nil || *snapshot.Form.Height != 420 {
		t.Fatalf("snapshot authored dimensions = %+v", snapshot.Form)
	}
	if snapshot.Form.Observed == nil || *snapshot.Form.Observed.Width != 530 || *snapshot.Form.Observed.Height != 450 {
		t.Fatalf("outer observations = %+v", snapshot.Form.Observed)
	}
	observedOnly := NormalizeFormSpec(FormSpec{Form: FormSpecForm{Name: "TestForm", Observed: snapshot.Form.Observed}})
	if observedOnly.Form.Width != nil || observedOnly.Form.Height != nil || observedOnly.Form.Caption != nil {
		t.Fatalf("observations became authoring intent: %+v", observedOnly.Form)
	}
}
