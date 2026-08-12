package hotspots

import (
	"encoding/json"
	"testing"
)

func TestReportJSONIncludesRawAndNormalizedSignals(t *testing.T) {
	report := BuildReport([]Input{{
		ID: "M.Run", Kind: "procedure", File: "m.bas", Module: "M", Name: "Run",
		RawSignals: map[string]int{"complexity": 2},
	}}, nil)
	data, err := json.Marshal(report)
	if err != nil {
		t.Fatal(err)
	}
	var decoded struct {
		SchemaVersion int      `json:"schema_version"`
		ScoreModel    string   `json:"score_model"`
		Procedures    []Entity `json:"procedures"`
	}
	if err := json.Unmarshal(data, &decoded); err != nil {
		t.Fatal(err)
	}
	if decoded.SchemaVersion != SchemaVersion || decoded.ScoreModel != ScoreModel {
		t.Fatalf("metadata = schema %d/model %q", decoded.SchemaVersion, decoded.ScoreModel)
	}
	if len(decoded.Procedures) != 1 {
		t.Fatalf("procedures = %#v", decoded.Procedures)
	}
	entity := decoded.Procedures[0]
	if entity.RawSignals["complexity"] != 2 {
		t.Fatalf("raw signals = %#v", entity.RawSignals)
	}
	if entity.NormalizedSignals == nil {
		t.Fatal("normalized signals were omitted")
	}
}
