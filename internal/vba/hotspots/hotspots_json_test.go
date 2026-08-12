package hotspots

import (
	"encoding/json"
	"strings"
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
	var raw map[string]any
	if err := json.Unmarshal(data, &raw); err != nil {
		t.Fatal(err)
	}
	procedures := raw["procedures"].([]any)
	if _, ok := procedures[0].(map[string]any)["selected_by"]; ok {
		t.Fatalf("unselected entity included selected_by: %#v", procedures[0])
	}
	selected := Select(report.Procedures, 1, 0)
	selectedReport := BuildReport([]Input{{
		ID: "M.Run", Kind: "procedure", File: "m.bas", Module: "M", Name: "Run",
		RawSignals: map[string]int{"complexity": 2},
	}}, nil)
	selectedReport.Procedures = selected
	selectedData, err := json.Marshal(selectedReport)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(selectedData), `"selected_by":{"top_n":true}`) {
		t.Fatalf("selected entity omitted selected_by: %s", selectedData)
	}
}
