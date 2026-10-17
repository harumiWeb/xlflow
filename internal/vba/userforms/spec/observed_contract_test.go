package spec

import (
	"bytes"
	"encoding/json"
	"reflect"
	"strings"
	"testing"
)

func TestNormalizeFormSpecDoesNotInventObservations(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			model := FormSpec{
				SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer",
				Form: FormSpecForm{Name: "Main", Caption: new(""), Width: new(240.0), Height: new(180.0)},
				Controls: []FormSpecControl{
					{ID: "text", Type: "TextBox", Name: "Text1", Text: new(""), Value: false, Left: new(0.0), TabIndex: new(0), Enabled: new(false), Visible: new(false)},
					{ID: "tabs", Type: "TabStrip", Name: "Tabs1", Tabs: []FormSpecTab{}, SelectedIndex: new(-1)},
				},
			}
			for range 3 {
				model = NormalizeFormSpec(model)
				body, err := MarshalSnapshot(format, model)
				if err != nil {
					t.Fatal(err)
				}
				if bytes.Contains(body, []byte("observed")) {
					t.Fatalf("normalization invented observations: %s", body)
				}
				model, err = ParseFormSpec(SpecInput{Format: format}, body)
				if err != nil {
					t.Fatal(err)
				}
				if len(model.ValidationWarnings) != 0 {
					t.Fatalf("normal authored data produced warnings: %+v", model.ValidationWarnings)
				}
			}
			text := model.Controls[0]
			if text.Text == nil || *text.Text != "" || text.Value != false || text.Left == nil || *text.Left != 0 || text.Enabled == nil || *text.Enabled || text.Visible == nil || *text.Visible || text.TabIndex == nil || *text.TabIndex != 0 || model.Controls[1].Tabs == nil {
				t.Fatalf("empty/false/zero values or empty tabs lost: %+v", model.Controls)
			}
		})
	}
}

func TestLegacyObservationsSurviveParsingWithoutPromotion(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			model := NormalizeFormSpec(FormSpec{
				SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer",
				Form:     FormSpecForm{Name: "Main", Observed: &FormSpecObservedForm{Caption: new("Captured"), Width: new(300.0), ClientWidth: new(280.0)}},
				Controls: []FormSpecControl{{ID: "text", Type: "TextBox", Name: "Text1", Text: new("Authored"), Observed: &FormSpecObservedControl{Text: new("Captured"), Left: new(0.0), Enabled: new(false), Value: false}}},
			})
			for range 3 {
				body, err := MarshalSnapshot(format, model)
				if err != nil {
					t.Fatal(err)
				}
				loaded, err := ParseFormSpec(SpecInput{Format: format}, body)
				if err != nil {
					t.Fatal(err)
				}
				if !reflect.DeepEqual(model.Form.Observed, loaded.Form.Observed) || !reflect.DeepEqual(model.Controls[0].Observed, loaded.Controls[0].Observed) {
					t.Fatalf("legacy observations changed: %s", body)
				}
				if loaded.Form.Caption != nil || loaded.Form.Width != nil || loaded.Controls[0].Left != nil || loaded.Controls[0].Enabled != nil || *loaded.Controls[0].Text != "Authored" || len(loaded.ValidationWarnings) != 0 {
					t.Fatalf("observations promoted or diagnosed: %+v", loaded)
				}
				model = loaded
			}
		})
	}
}

func TestCapturedWarningsReplaceSnapshotContainerWarnings(t *testing.T) {
	const source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","observed":{"insideWidth":200}},"controls":[{"id":"label","type":"Label","name":"Label1","unsupported":["specialEffect"],"observed":{"caption":"old","width":10}}],"warnings":[{"code":"unsupported_properties","message":"Unsupported properties remain in the binary model.","control":"Label1"},{"code":"compatibility_artifact_unsynchronized","message":"Compatibility artifacts are not synchronized."}]}`
	var model FormSpec
	if err := json.Unmarshal([]byte(source), &model); err != nil {
		t.Fatal(err)
	}
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			body, err := MarshalSnapshot(format, model)
			if err != nil {
				t.Fatal(err)
			}
			loaded, err := ParseFormSpec(SpecInput{Format: format}, body)
			if err != nil {
				t.Fatal(err)
			}
			issues := loaded.ValidationWarnings
			if len(issues) != 2 || issues[0].Code != model.Warnings[0].Code || issues[1].Code != model.Warnings[1].Code || issues[0].Field != "warnings[0].message" || issues[1].Field != "warnings[1].message" || issues[0].Message != model.Warnings[0].Message {
				t.Fatalf("captured warnings = %+v", issues)
			}
			for _, issue := range issues {
				if issue.Severity != SeverityWarning || issue.Code == "UFV013" {
					t.Fatalf("snapshot container generated a support warning: %+v", issue)
				}
			}
		})
	}
}

func TestCapturedWarningEntriesAreValidated(t *testing.T) {
	for _, entry := range []string{`false`, `{"message":123}`, `{"message":"x","code":123}`, `{"message":"x","control":false}`} {
		body := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[],"warnings":[` + entry + `]}`)
		issues, err := ValidateFormSpecSource(SpecInput{Format: "json"}, body)
		if err != nil {
			t.Fatal(err)
		}
		if !hasValidationErrors(issues) {
			t.Fatalf("accepted malformed warning %s: %+v", entry, issues)
		}
		for _, issue := range issues {
			if issue.Field != "warnings" && !strings.HasPrefix(issue.Field, "warnings[0]") {
				t.Fatalf("unexpected diagnostic field: %+v", issue)
			}
		}
	}
}
