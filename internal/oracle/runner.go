package oracle

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
)

const (
	CommandName                  = "vbe-oracle"
	DefaultTimeout               = 5 * time.Minute
	bridgeShutdownMargin         = 30 * time.Second
	OutcomeAccepted              = "accepted"
	OutcomeRejected              = "rejected"
	OutcomeInfrastructureFailure = "infrastructure_failure"
)

type BridgeExecutor interface {
	Execute(context.Context, excelbridge.Request) (excelbridge.Response, error)
}

type Options struct {
	ManifestPath      string
	CaseIDs           []string
	Strict            bool
	PromoteObserved   bool
	DiagnosticMeaning map[string]string
	Timeout           time.Duration
	Executor          BridgeExecutor
	Now               func() time.Time
}

type Report struct {
	SchemaVersion int          `json:"schema_version"`
	Status        string       `json:"status"`
	Outcome       string       `json:"outcome,omitempty"`
	Cases         []CaseResult `json:"cases,omitempty"`
	Controls      []CaseResult `json:"controls,omitempty"`
	Error         *ReportError `json:"error,omitempty"`
}

type ReportError struct {
	Code    string `json:"code"`
	Message string `json:"message"`
}

type CaseResult struct {
	ID               string               `json:"id"`
	Expected         string               `json:"expected,omitempty"`
	Outcome          string               `json:"outcome"`
	EvidencePhase    string               `json:"evidence_phase,omitempty"`
	LastStage        string               `json:"last_stage,omitempty"`
	CompileInvoked   bool                 `json:"compile_invoked,omitempty"`
	CleanupConfirmed bool                 `json:"cleanup_confirmed,omitempty"`
	DurationMS       int64                `json:"duration_ms,omitempty"`
	Dialog           any                  `json:"dialog,omitempty"`
	Location         any                  `json:"location,omitempty"`
	ExcelProcessID   int                  `json:"excel_process_id,omitempty"`
	Metadata         VerificationMetadata `json:"metadata,omitempty"`
	Error            string               `json:"error,omitempty"`
}

type ExitError struct {
	Code    int
	Message string
	Report  *Report
}

func (e *ExitError) Error() string {
	if e == nil {
		return ""
	}
	return e.Message
}

type fixtureRequest struct {
	SchemaVersion int             `json:"schema_version"`
	CaseID        string          `json:"case_id"`
	ProbeMode     string          `json:"probe_mode"`
	Visible       bool            `json:"visible"`
	TimeoutMS     int64           `json:"timeout_ms"`
	Modules       []fixtureModule `json:"modules"`
}

type fixtureModule struct {
	Name       string `json:"name"`
	Kind       string `json:"kind"`
	SourcePath string `json:"source_path"`
}

type bridgeObservation struct {
	Outcome          string               `json:"outcome"`
	EvidencePhase    string               `json:"evidence_phase"`
	LastStage        string               `json:"last_stage"`
	DurationMS       int64                `json:"duration_ms"`
	CompileInvoked   bool                 `json:"compile_invoked"`
	CleanupConfirmed bool                 `json:"cleanup_confirmed"`
	Dialog           any                  `json:"dialog,omitempty"`
	Location         any                  `json:"location,omitempty"`
	ExcelProcessID   int                  `json:"excel_process_id,omitempty"`
	Metadata         VerificationMetadata `json:"metadata,omitempty"`
	Excel            map[string]any       `json:"excel,omitempty"`
	Error            string               `json:"error,omitempty"`
	ErrorCode        string               `json:"error_code,omitempty"`
}

// Run loads, validates, and executes the selected cases. It intentionally has
// no parallelism: an Excel/VBE instance is a single-user UI resource.
func Run(ctx context.Context, opts Options) (Report, error) {
	report := Report{SchemaVersion: SchemaVersion, Status: "ok"}
	if runtime.GOOS != "windows" {
		return failReport(&report, 3, "oracle is available only on Windows with Excel installed", "unsupported_host")
	}
	return runValidated(ctx, opts)
}

// runValidated contains the deterministic orchestration after the Windows
// host gate. Tests can inject a fake bridge here without requiring Excel.
func runValidated(ctx context.Context, opts Options) (Report, error) {
	report := Report{SchemaVersion: SchemaVersion, Status: "ok"}
	if opts.ManifestPath == "" {
		opts.ManifestPath = filepath.Join("testdata", "vbe-oracle", "manifest.json")
	}
	manifest, root, err := LoadManifest(opts.ManifestPath)
	if err != nil {
		return failReport(&report, 2, err.Error(), "fixture_invalid")
	}
	if opts.Timeout < 0 {
		return failReport(&report, 2, "oracle timeout must be greater than zero", "timeout_invalid")
	}
	if opts.Timeout == 0 {
		opts.Timeout = DefaultTimeout
	}
	if opts.Now == nil {
		opts.Now = time.Now
	}
	if opts.Executor == nil {
		provider := excelbridge.DotNetProvider{RootDir: root, PreferRepoLocal: true}
		opts.Executor = provider
	}
	entries := make(map[string]ManifestEntry, len(manifest.Cases))
	for _, e := range manifest.Cases {
		entries[e.ID] = e
	}
	selected, err := selectEntries(manifest, opts.CaseIDs)
	if err != nil {
		return failReport(&report, 2, err.Error(), "case_selection_invalid")
	}
	if opts.PromoteObserved && len(opts.CaseIDs) == 0 {
		return failReport(&report, 2, "--promote-observed requires one or more explicit --case values", "promotion_invalid")
	}
	// Controls always run first, even when --case selects another fixture.
	controlIDs := []string{manifest.Controls.Accept, manifest.Controls.Reject}
	seen := map[string]bool{}
	for _, id := range controlIDs {
		if seen[id] {
			continue
		}
		seen[id] = true
		result, runErr := runEntry(ctx, root, entries[id], opts.Timeout, opts.Executor)
		report.Controls = append(report.Controls, result)
		if runErr != nil {
			return failRun(&report, runErr)
		}
		expected := ExpectedAccepted
		if id == manifest.Controls.Reject {
			expected = ExpectedRejected
		}
		if result.Outcome != expected {
			return failReport(&report, 3, fmt.Sprintf("oracle control %q returned %s, want %s", id, result.Outcome, expected), "control_failed")
		}
	}
	for _, entry := range selected {
		if seen[entry.ID] {
			continue
		}
		seen[entry.ID] = true
		result, runErr := runEntry(ctx, root, entry, opts.Timeout, opts.Executor)
		report.Cases = append(report.Cases, result)
		if runErr != nil {
			return failRun(&report, runErr)
		}
		if opts.Strict && result.Outcome != result.Expected {
			return failReport(&report, 1, fmt.Sprintf("oracle case %q expected %s, observed %s", entry.ID, result.Expected, result.Outcome), "expectation_mismatch")
		}
	}
	if opts.PromoteObserved {
		promotionResults := append(append([]CaseResult(nil), report.Controls...), report.Cases...)
		if err := promoteSelected(root, selected, promotionResults, opts.DiagnosticMeaning, opts.Now()); err != nil {
			return failReport(&report, 2, err.Error(), "promotion_invalid")
		}
	}
	return report, nil
}

func selectEntries(manifest Manifest, ids []string) ([]ManifestEntry, error) {
	if len(ids) == 0 {
		return append([]ManifestEntry(nil), manifest.Cases...), nil
	}
	entries := make(map[string]ManifestEntry, len(manifest.Cases))
	for _, e := range manifest.Cases {
		entries[e.ID] = e
	}
	seen := map[string]bool{}
	selected := make([]ManifestEntry, 0, len(ids))
	for _, id := range ids {
		id = strings.TrimSpace(id)
		if id == "" {
			return nil, errors.New("--case cannot be empty")
		}
		if seen[id] {
			return nil, fmt.Errorf("duplicate --case %q", id)
		}
		seen[id] = true
		e, ok := entries[id]
		if !ok {
			return nil, fmt.Errorf("unknown oracle case %q", id)
		}
		selected = append(selected, e)
	}
	return selected, nil
}

// SelectCases resolves an explicit filter while preserving manifest order.
// An empty filter returns every manifest case in that order.
func SelectCases(manifest Manifest, ids []string) ([]ManifestEntry, error) {
	return selectEntries(manifest, ids)
}

func runEntry(parent context.Context, root string, entry ManifestEntry, timeout time.Duration, executor BridgeExecutor) (CaseResult, error) {
	started := time.Now()
	c, _, err := LoadCase(root, entry)
	if err != nil {
		return CaseResult{ID: entry.ID, Outcome: OutcomeInfrastructureFailure, Error: err.Error()}, &ExitError{Code: 2, Message: err.Error()}
	}
	modules := make([]fixtureModule, 0, len(c.Modules))
	casePath := entry.Path
	if casePath == "" {
		casePath = filepath.Join("cases", entry.ID, "case.json")
	}
	caseDir, err := confinedPath(root, filepath.Dir(casePath))
	if err != nil {
		return CaseResult{ID: c.ID, Outcome: OutcomeInfrastructureFailure, Error: err.Error()}, &ExitError{Code: 2, Message: err.Error()}
	}
	for _, module := range c.Modules {
		modulePath, err := confinedPath(caseDir, module.Path)
		if err != nil {
			return CaseResult{ID: c.ID, Outcome: OutcomeInfrastructureFailure, Error: err.Error()}, &ExitError{Code: 2, Message: err.Error()}
		}
		modules = append(modules, fixtureModule{Name: module.Name, Kind: module.Kind, SourcePath: modulePath})
	}
	payload, err := json.Marshal(fixtureRequest{SchemaVersion: SchemaVersion, CaseID: c.ID, ProbeMode: c.Probe.Mode, TimeoutMS: timeout.Milliseconds(), Modules: modules})
	if err != nil {
		return CaseResult{ID: c.ID, Outcome: OutcomeInfrastructureFailure, Error: err.Error()}, &ExitError{Code: 3, Message: err.Error()}
	}
	// Let the bridge timeout expire first so it can close Excel and report
	// structured cleanup status before the Go process deadline fires.
	ctx, cancel := context.WithTimeout(parent, timeout+bridgeShutdownMargin)
	defer cancel()
	response, execErr := executor.Execute(ctx, excelbridge.Request{Command: CommandName, Args: map[string]string{
		"PlanJson64": base64.StdEncoding.EncodeToString(payload),
		"TimeoutMs":  fmt.Sprintf("%d", timeout.Milliseconds()),
	}})
	if execErr != nil || response.TimedOut {
		msg := "oracle bridge execution failed"
		if execErr != nil {
			msg = execErr.Error()
		}
		return CaseResult{ID: c.ID, Outcome: OutcomeInfrastructureFailure, Error: msg, DurationMS: time.Since(started).Milliseconds()}, &ExitError{Code: 3, Message: msg}
	}
	obs, err := decodeObservation(response.Stdout)
	if err != nil {
		code := 3
		if strings.Contains(strings.ToLower(bridgeErrorCode(response.Stdout)), "plan_invalid") {
			code = 2
		}
		return CaseResult{ID: c.ID, Outcome: OutcomeInfrastructureFailure, Error: err.Error()}, &ExitError{Code: code, Message: err.Error()}
	}
	metadata := obs.Metadata
	if metadata.ExcelVersion == "" {
		metadata.ExcelVersion = stringValue(obs.Excel, "version")
	}
	if metadata.ExcelBuild == "" {
		metadata.ExcelBuild = stringValue(obs.Excel, "build")
	}
	if metadata.Bitness == "" {
		metadata.Bitness = stringValue(obs.Excel, "bitness")
	}
	if metadata.Locale == "" {
		metadata.Locale = stringValue(obs.Excel, "locale")
	}
	result := CaseResult{ID: c.ID, Expected: c.VBE.Expected, Outcome: obs.Outcome, EvidencePhase: obs.EvidencePhase, LastStage: obs.LastStage, CompileInvoked: obs.CompileInvoked, CleanupConfirmed: obs.CleanupConfirmed, DurationMS: obs.DurationMS, Dialog: obs.Dialog, Location: obs.Location, ExcelProcessID: obs.ExcelProcessID, Metadata: metadata, Error: obs.Error}
	if result.DurationMS == 0 {
		result.DurationMS = time.Since(started).Milliseconds()
	}
	if result.Outcome != OutcomeAccepted && result.Outcome != OutcomeRejected {
		result.Outcome = OutcomeInfrastructureFailure
	}
	if result.Outcome == OutcomeInfrastructureFailure || !result.CleanupConfirmed {
		result.Outcome = OutcomeInfrastructureFailure
		if result.Error == "" {
			result.Error = "oracle cleanup or infrastructure failure"
		}
		code := 3
		if strings.Contains(strings.ToLower(obs.ErrorCode), "plan_invalid") || strings.Contains(strings.ToLower(obs.ErrorCode), "args_invalid") {
			code = 2
		}
		return result, &ExitError{Code: code, Message: result.Error}
	}
	return result, nil
}

func decodeObservation(data []byte) (bridgeObservation, error) {
	var raw map[string]json.RawMessage
	if err := json.Unmarshal(data, &raw); err != nil {
		return bridgeObservation{}, fmt.Errorf("malformed oracle bridge output: %w", err)
	}
	if oracleJSON, ok := raw["oracle"]; ok {
		var observation bridgeObservation
		if err := json.Unmarshal(oracleJSON, &observation); err != nil {
			return bridgeObservation{}, fmt.Errorf("malformed oracle observation: %w", err)
		}
		return observation, nil
	}
	var failed struct {
		Error struct {
			Details struct {
				Oracle *bridgeObservation `json:"oracle"`
			} `json:"details"`
		} `json:"error"`
	}
	if err := json.Unmarshal(data, &failed); err == nil && failed.Error.Details.Oracle != nil {
		return *failed.Error.Details.Oracle, nil
	}
	if errorJSON, ok := raw["error"]; ok {
		var bridgeErr struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		}
		_ = json.Unmarshal(errorJSON, &bridgeErr)
		message := bridgeErr.Message
		if message == "" {
			message = "oracle bridge returned a failure"
		}
		return bridgeObservation{Outcome: OutcomeInfrastructureFailure, EvidencePhase: EvidenceUnknown, CleanupConfirmed: false, Error: bridgeErr.Code + ": " + message, ErrorCode: bridgeErr.Code}, nil
	}
	var direct bridgeObservation
	if err := json.Unmarshal(data, &direct); err != nil {
		return bridgeObservation{}, fmt.Errorf("malformed oracle bridge output: %w", err)
	}
	if direct.Outcome == "" {
		return bridgeObservation{}, errors.New("malformed oracle bridge output: missing outcome")
	}
	return direct, nil
}

func bridgeErrorCode(data []byte) string {
	var value struct {
		Error struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if json.Unmarshal(data, &value) == nil {
		return value.Error.Code
	}
	return ""
}

func stringValue(values map[string]any, key string) string {
	if value, ok := values[key]; ok {
		if text, ok := value.(string); ok {
			return text
		}
	}
	return ""
}

func failReport(report *Report, code int, message, kind string) (Report, error) {
	report.Status, report.Outcome = "failed", OutcomeInfrastructureFailure
	report.Error = &ReportError{Code: kind, Message: message}
	return *report, &ExitError{Code: code, Message: message, Report: report}
}

func failRun(report *Report, err error) (Report, error) {
	var exit *ExitError
	if errors.As(err, &exit) {
		report.Status, report.Outcome = "failed", OutcomeInfrastructureFailure
		report.Error = &ReportError{Code: "oracle_failure", Message: exit.Message}
		exit.Report = report
		return *report, exit
	}
	return failReport(report, 3, err.Error(), "oracle_failure")
}

type promotionUpdate struct {
	path     string
	original []byte
	body     []byte
}

type promotionStage struct {
	promotionUpdate
	temp string
}

func promoteSelected(root string, selected []ManifestEntry, results []CaseResult, meanings map[string]string, now time.Time) error {
	resultByID := map[string]CaseResult{}
	for _, result := range results {
		resultByID[result.ID] = result
	}
	updates := []promotionUpdate{}
	for _, entry := range selected {
		casePath := entry.Path
		if casePath == "" {
			casePath = filepath.Join("cases", entry.ID, "case.json")
		}
		path, err := confinedPath(root, casePath)
		if err != nil {
			return err
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		var c Case
		if err := json.Unmarshal(body, &c); err != nil {
			return err
		}
		if c.VBE.Expected != ExpectedObserve {
			return fmt.Errorf("oracle case %q is already asserted", c.ID)
		}
		result, ok := resultByID[c.ID]
		if !ok {
			return fmt.Errorf("oracle case %q was not executed", c.ID)
		}
		if result.Outcome != OutcomeAccepted && result.Outcome != OutcomeRejected {
			return fmt.Errorf("oracle case %q has non-evidence outcome %s", c.ID, result.Outcome)
		}
		meaning := MeaningCompileError
		if result.Outcome == OutcomeAccepted {
			meaning = strings.TrimSpace(meanings[c.ID])
			if meaning == "" || (meaning != MeaningSpecification && meaning != MeaningPolicy && meaning != MeaningMaintainability) {
				return fmt.Errorf("accepted case %q requires diagnostic meaning specification, policy, or maintainability", c.ID)
			}
		}
		c.VBE.Expected, c.VBE.EvidencePhase, c.VBE.DiagnosticMeaning = result.Outcome, EvidenceCompile, meaning
		c.Provenance.Status = "asserted"
		metadata := result.Metadata
		if missingMetadata(metadata.ExcelVersion) || missingMetadata(metadata.ExcelBuild) || missingMetadata(metadata.Bitness) || missingMetadata(metadata.Locale) {
			return fmt.Errorf("oracle case %q cannot be promoted without complete Excel version, build, bitness, and locale metadata", c.ID)
		}
		metadata.VerifiedAt = now.UTC().Format(time.RFC3339)
		c.Provenance.VerifiedOn = append(c.Provenance.VerifiedOn, metadata)
		updated, err := json.MarshalIndent(c, "", "  ")
		if err != nil {
			return err
		}
		updated = append(updated, '\n')
		updates = append(updates, promotionUpdate{path: path, original: body, body: updated})
	}
	// Stage every new document before touching an existing fixture. Validation
	// above therefore cannot leave a partially promoted batch behind.
	staged := make([]promotionStage, 0, len(updates))
	for _, u := range updates {
		tmp, err := os.CreateTemp(filepath.Dir(u.path), ".oracle-promote-*")
		if err != nil {
			for _, item := range staged {
				_ = os.Remove(item.temp)
			}
			return err
		}
		name := tmp.Name()
		if _, err = tmp.Write(u.body); err == nil {
			err = tmp.Close()
		} else {
			_ = tmp.Close()
		}
		if err != nil {
			_ = os.Remove(name)
			for _, item := range staged {
				_ = os.Remove(item.temp)
			}
			return err
		}
		staged = append(staged, promotionStage{promotionUpdate: u, temp: name})
	}
	committed := 0
	for i, item := range staged {
		if err := os.Rename(item.temp, item.path); err != nil {
			for _, left := range staged[i:] {
				_ = os.Remove(left.temp)
			}
			rollbackPromotion(staged[:committed])
			return err
		}
		committed++
	}
	return nil
}

func missingMetadata(value string) bool {
	value = strings.TrimSpace(value)
	return value == "" || strings.EqualFold(value, "unknown")
}

// PromoteObserved atomically promotes selected observe fixtures after the
// caller has completed a successful oracle run.
func PromoteObserved(root string, selected []ManifestEntry, results []CaseResult, meanings map[string]string, now time.Time) error {
	if len(selected) == 0 {
		return errors.New("promotion requires explicit case selection")
	}
	return promoteSelected(root, selected, results, meanings, now)
}

func rollbackPromotion(committed []promotionStage) {
	for i := len(committed) - 1; i >= 0; i-- {
		item := committed[i]
		_ = os.WriteFile(item.path, item.original, 0o644)
	}
}
