package oracle

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"testing"
	"time"

	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
)

func TestNormalizeDiagnosticsAndContract(t *testing.T) {
	expect := AnalysisExpectation{ExpectedDiagnostics: []DiagnosticExpectation{{Code: "VBA001", Severity: "warning", Range: &Range{StartLine: 1, EndLine: 1}}}, ForbiddenDiagnostics: []DiagnosticExpectation{{Code: "VBA999"}}}
	actual := []Diagnostic{{Code: "VBA001", Severity: "warning", Range: &Range{StartLine: 1, EndLine: 1}, Surface: "lint"}}
	if err := CheckDiagnosticContract(expect, actual); err != nil {
		t.Fatal(err)
	}
	if err := CheckDiagnosticContract(expect, append(actual, Diagnostic{Code: "VBA999"})); err == nil {
		t.Fatal("expected forbidden diagnostic failure")
	}
	expect.ExpectedDiagnostics = append(expect.ExpectedDiagnostics, expect.ExpectedDiagnostics[0])
	if err := CheckDiagnosticContract(expect, actual); err == nil {
		t.Fatal("expected duplicate expectation failure")
	}
}

func TestDecodeObservationSupportsBridgeExtension(t *testing.T) {
	data := []byte(`{"oracle":{"outcome":"accepted","evidence_phase":"compile","cleanup_confirmed":true}}`)
	obs, err := decodeObservation(data)
	if err != nil || obs.Outcome != OutcomeAccepted || !obs.CleanupConfirmed {
		t.Fatalf("obs=%+v err=%v", obs, err)
	}
	data = []byte(`{"error":{"details":{"oracle":{"outcome":"infrastructure_failure","cleanup_confirmed":false}}}}`)
	obs, err = decodeObservation(data)
	if err != nil || obs.Outcome != OutcomeInfrastructureFailure {
		t.Fatalf("failed obs=%+v err=%v", obs, err)
	}
}

func TestFixtureRequestUsesBridgePlanContract(t *testing.T) {
	payload, err := json.Marshal(fixtureRequest{SchemaVersion: SchemaVersion, CaseID: "x", ProbeMode: ProbeCompile, Modules: []fixtureModule{{Name: "Main", Kind: "standard", SourcePath: "C:/x/Main.bas"}}})
	if err != nil {
		t.Fatal(err)
	}
	encoded := base64.StdEncoding.EncodeToString(payload)
	decoded, err := base64.StdEncoding.DecodeString(encoded)
	if err != nil || !json.Valid(decoded) {
		t.Fatalf("payload is not valid base64 JSON: %v", err)
	}
	var plan fixtureRequest
	if err := json.Unmarshal(decoded, &plan); err != nil || plan.SchemaVersion != SchemaVersion || plan.CaseID != "x" || plan.ProbeMode != ProbeCompile {
		t.Fatalf("plan=%+v err=%v", plan, err)
	}
}

type oracleFakeBridge struct{ calls []string }

func (f *oracleFakeBridge) Execute(_ context.Context, req excelbridge.Request) (excelbridge.Response, error) {
	planBytes, err := base64.StdEncoding.DecodeString(req.Args["PlanJson64"])
	if err != nil {
		return excelbridge.Response{}, err
	}
	var plan fixtureRequest
	if err := json.Unmarshal(planBytes, &plan); err != nil {
		return excelbridge.Response{}, err
	}
	f.calls = append(f.calls, plan.CaseID)
	outcome := OutcomeAccepted
	if plan.CaseID == "sample-reject" {
		outcome = OutcomeRejected
	}
	body, _ := json.Marshal(map[string]any{"oracle": map[string]any{"outcome": outcome, "evidence_phase": EvidenceCompile, "cleanup_confirmed": true, "excel": map[string]string{"version": "16.0", "build": "17932", "bitness": "x64", "locale": "ja-JP"}}})
	return excelbridge.Response{Stdout: body}, nil
}

func TestRunControlsStrictAndPromotion(t *testing.T) {
	manifestPath, _ := writeFixture(t, ExpectedObserve)
	bridge := &oracleFakeBridge{}
	report, err := Run(context.Background(), Options{ManifestPath: manifestPath, CaseIDs: []string{"sample"}, PromoteObserved: true, DiagnosticMeaning: map[string]string{"sample": MeaningSpecification}, Executor: bridge, Timeout: time.Second})
	if err != nil {
		t.Fatalf("run promotion failed: %v report=%+v", err, report)
	}
	if len(bridge.calls) != 2 || bridge.calls[0] != "sample" || bridge.calls[1] != "sample-reject" {
		t.Fatalf("calls=%v", bridge.calls)
	}
	_, root, _ := LoadManifest(manifestPath)
	c, _, err := LoadCase(root, ManifestEntry{ID: "sample"})
	if err != nil {
		t.Fatal(err)
	}
	if c.VBE.Expected != ExpectedAccepted || c.Provenance.Status != "asserted" {
		t.Fatalf("promotion=%+v", c)
	}
}

func TestPromoteObservedRequiresExcelMetadata(t *testing.T) {
	manifestPath, _ := writeFixture(t, ExpectedObserve)
	_, root, err := LoadManifest(manifestPath)
	if err != nil {
		t.Fatal(err)
	}
	err = PromoteObserved(root, []ManifestEntry{{ID: "sample"}}, []CaseResult{{ID: "sample", Outcome: OutcomeAccepted}}, map[string]string{"sample": MeaningSpecification}, time.Now())
	if err == nil {
		t.Fatal("expected incomplete metadata promotion failure")
	}
}

func TestRunRejectsImplicitPromotionBeforeBridgeExecution(t *testing.T) {
	manifestPath, _ := writeFixture(t, ExpectedObserve)
	bridge := &oracleFakeBridge{}
	_, err := Run(context.Background(), Options{ManifestPath: manifestPath, PromoteObserved: true, Executor: bridge, Timeout: time.Second})
	if err == nil {
		t.Fatal("expected implicit promotion validation error")
	}
	if len(bridge.calls) != 0 {
		t.Fatalf("bridge calls=%v, want none", bridge.calls)
	}
	var exit *ExitError
	if !errors.As(err, &exit) || exit.Code != 2 {
		t.Fatalf("err=%v, want exit code 2", err)
	}
}
