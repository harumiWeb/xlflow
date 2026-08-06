package oracle

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"
)

const (
	SchemaVersion = 1

	ProbeCompile = "compile"

	ExpectedObserve  = "observe"
	ExpectedAccepted = "accepted"
	ExpectedRejected = "rejected"

	EvidenceUnknown = "unknown"
	EvidenceCompile = "compile"

	MeaningObservation     = "observation"
	MeaningCompileError    = "compile-error"
	MeaningRuntimeError    = "runtime-error"
	MeaningSpecification   = "specification"
	MeaningPolicy          = "policy"
	MeaningMaintainability = "maintainability"
)

var caseIDPattern = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._-]*$`)
var vbNamePattern = regexp.MustCompile(`(?im)^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"`)

type Manifest struct {
	SchemaVersion int             `json:"schema_version"`
	Cases         []ManifestEntry `json:"cases"`
	Controls      Controls        `json:"controls"`
}

func (m *Manifest) UnmarshalJSON(data []byte) error {
	var value struct {
		SchemaVersion int             `json:"schema_version"`
		Cases         []ManifestEntry `json:"cases"`
		Controls      Controls        `json:"controls"`
		KnownAccept   string          `json:"known_accept"`
		KnownReject   string          `json:"known_reject"`
	}
	if err := json.Unmarshal(data, &value); err != nil {
		return err
	}
	m.SchemaVersion, m.Cases, m.Controls = value.SchemaVersion, value.Cases, value.Controls
	if m.Controls.Accept == "" {
		m.Controls.Accept = value.KnownAccept
	}
	if m.Controls.Reject == "" {
		m.Controls.Reject = value.KnownReject
	}
	return nil
}

// ManifestEntry keeps the manifest extensible while allowing the concise
// `"cases": ["case-id"]` form used by early local manifests.
type ManifestEntry struct {
	ID          string `json:"id"`
	Path        string `json:"path,omitempty"`
	ControlRole string `json:"control_role,omitempty"`
	Role        string `json:"role,omitempty"`
}

func (e *ManifestEntry) UnmarshalJSON(data []byte) error {
	var id string
	if err := json.Unmarshal(data, &id); err == nil {
		e.ID = strings.TrimSpace(id)
		return nil
	}
	type alias ManifestEntry
	var value alias
	if err := json.Unmarshal(data, &value); err != nil {
		return err
	}
	*e = ManifestEntry(value)
	if e.ControlRole == "" {
		e.ControlRole = e.Role
	}
	return nil
}

type Controls struct {
	Accept string `json:"accept"`
	Reject string `json:"reject"`
}

func (c *Controls) UnmarshalJSON(data []byte) error {
	type alias Controls
	var value alias
	if err := json.Unmarshal(data, &value); err == nil && (value.Accept != "" || value.Reject != "") {
		*c = Controls(value)
		return nil
	}
	var legacy struct {
		KnownAccept string `json:"known_accept"`
		KnownReject string `json:"known_reject"`
	}
	if err := json.Unmarshal(data, &legacy); err != nil {
		return err
	}
	c.Accept, c.Reject = legacy.KnownAccept, legacy.KnownReject
	return nil
}

type Case struct {
	SchemaVersion int                 `json:"schema_version"`
	ID            string              `json:"id"`
	Description   string              `json:"description,omitempty"`
	Modules       []Module            `json:"modules"`
	Probe         Probe               `json:"probe"`
	VBE           VBEExpectation      `json:"vbe"`
	Analysis      AnalysisExpectation `json:"analysis,omitempty"`
	Provenance    Provenance          `json:"provenance"`
}

type Module struct {
	Name  string `json:"name"`
	Kind  string `json:"kind"`
	Path  string `json:"path"`
	Entry bool   `json:"entry,omitempty"`
}

type Probe struct {
	Mode string `json:"mode"`
}

type VBEExpectation struct {
	Expected          string `json:"expected"`
	EvidencePhase     string `json:"evidence_phase"`
	DiagnosticMeaning string `json:"diagnostic_meaning"`
}

type AnalysisExpectation struct {
	ExpectedDiagnostics  []DiagnosticExpectation `json:"expected_diagnostics,omitempty"`
	ForbiddenDiagnostics []DiagnosticExpectation `json:"forbidden_diagnostics,omitempty"`
}

type DiagnosticExpectation struct {
	Code     string   `json:"code"`
	Severity string   `json:"severity,omitempty"`
	Range    *Range   `json:"range,omitempty"`
	Surfaces []string `json:"surfaces,omitempty"`
}

type Range struct {
	StartLine   int `json:"start_line,omitempty"`
	StartColumn int `json:"start_column,omitempty"`
	EndLine     int `json:"end_line,omitempty"`
	EndColumn   int `json:"end_column,omitempty"`
}

type Provenance struct {
	Status     string                 `json:"status"`
	VerifiedOn []VerificationMetadata `json:"verified_on,omitempty"`
}

type VerificationMetadata struct {
	ExcelVersion string `json:"excel_version,omitempty"`
	ExcelBuild   string `json:"excel_build,omitempty"`
	Bitness      string `json:"bitness,omitempty"`
	Locale       string `json:"locale,omitempty"`
	VerifiedAt   string `json:"verified_at,omitempty"`
}

func (v *VerificationMetadata) UnmarshalJSON(data []byte) error {
	var value struct {
		ExcelVersion string `json:"excel_version,omitempty"`
		ExcelBuild   string `json:"excel_build,omitempty"`
		Bitness      string `json:"bitness,omitempty"`
		Locale       string `json:"locale,omitempty"`
		VerifiedAt   string `json:"verified_at,omitempty"`
	}
	if err := json.Unmarshal(data, &value); err == nil {
		*v = VerificationMetadata{ExcelVersion: value.ExcelVersion, ExcelBuild: value.ExcelBuild, Bitness: value.Bitness, Locale: value.Locale, VerifiedAt: value.VerifiedAt}
		return nil
	}
	var timestamp string
	if err := json.Unmarshal(data, &timestamp); err != nil {
		return err
	}
	v.VerifiedAt = timestamp
	return nil
}

func LoadManifest(path string) (Manifest, string, error) {
	path, err := filepath.Abs(path)
	if err != nil {
		return Manifest{}, "", err
	}
	body, err := os.ReadFile(path)
	if err != nil {
		return Manifest{}, "", fmt.Errorf("read oracle manifest: %w", err)
	}
	var manifest Manifest
	if err := json.Unmarshal(body, &manifest); err != nil {
		return Manifest{}, "", fmt.Errorf("decode oracle manifest: %w", err)
	}
	root := filepath.Dir(path)
	if err := ValidateManifest(manifest); err != nil {
		return Manifest{}, "", err
	}
	return manifest, root, nil
}

func ValidateManifest(manifest Manifest) error {
	if manifest.SchemaVersion != SchemaVersion {
		return fmt.Errorf("unsupported oracle manifest schema_version %d (want %d)", manifest.SchemaVersion, SchemaVersion)
	}
	if len(manifest.Cases) == 0 {
		return errors.New("oracle manifest contains no cases")
	}
	seen := map[string]struct{}{}
	for _, entry := range manifest.Cases {
		if !caseIDPattern.MatchString(entry.ID) {
			return fmt.Errorf("invalid oracle case id %q", entry.ID)
		}
		if _, ok := seen[entry.ID]; ok {
			return fmt.Errorf("duplicate oracle case %q", entry.ID)
		}
		if entry.Path != "" {
			if filepath.IsAbs(entry.Path) || !strings.EqualFold(filepath.Clean(filepath.FromSlash(entry.Path)), filepath.Join("cases", entry.ID, "case.json")) {
				return fmt.Errorf("oracle case %q path must be cases/%s/case.json", entry.ID, entry.ID)
			}
		}
		seen[entry.ID] = struct{}{}
		role := strings.ToLower(strings.TrimSpace(entry.ControlRole))
		if role != "" && role != "accept" && role != "reject" {
			return fmt.Errorf("case %q has invalid control_role %q", entry.ID, entry.ControlRole)
		}
		if role == "accept" && manifest.Controls.Accept != "" && entry.ID != manifest.Controls.Accept {
			return fmt.Errorf("case %q is marked control_role=accept but controls.accept is %q", entry.ID, manifest.Controls.Accept)
		}
		if role == "reject" && manifest.Controls.Reject != "" && entry.ID != manifest.Controls.Reject {
			return fmt.Errorf("case %q is marked control_role=reject but controls.reject is %q", entry.ID, manifest.Controls.Reject)
		}
	}
	if manifest.Controls.Accept == "" || manifest.Controls.Reject == "" {
		return errors.New("oracle manifest must declare controls.accept and controls.reject")
	}
	if _, ok := seen[manifest.Controls.Accept]; !ok {
		return fmt.Errorf("controls.accept references unknown case %q", manifest.Controls.Accept)
	}
	if _, ok := seen[manifest.Controls.Reject]; !ok {
		return fmt.Errorf("controls.reject references unknown case %q", manifest.Controls.Reject)
	}
	if manifest.Controls.Accept == manifest.Controls.Reject {
		return errors.New("controls.accept and controls.reject must differ")
	}
	return nil
}

func LoadCase(manifestRoot string, entry ManifestEntry) (Case, map[string][]byte, error) {
	casePath := entry.Path
	if casePath == "" {
		casePath = filepath.Join("cases", entry.ID, "case.json")
	}
	path, err := confinedPath(manifestRoot, casePath)
	if err != nil {
		return Case{}, nil, err
	}
	body, err := os.ReadFile(path)
	if err != nil {
		return Case{}, nil, fmt.Errorf("read oracle case %q: %w", entry.ID, err)
	}
	var c Case
	if err := json.Unmarshal(body, &c); err != nil {
		return Case{}, nil, fmt.Errorf("decode oracle case %q: %w", entry.ID, err)
	}
	caseDir := filepath.Dir(path)
	if !strings.EqualFold(filepath.Base(path), "case.json") || !strings.EqualFold(filepath.Base(caseDir), entry.ID) || !strings.EqualFold(filepath.Base(filepath.Dir(caseDir)), "cases") {
		return Case{}, nil, fmt.Errorf("oracle case %q must be stored at cases/%s/case.json", entry.ID, entry.ID)
	}
	if err := ValidateCase(c, entry.ID, caseDir); err != nil {
		return Case{}, nil, err
	}
	sources := make(map[string][]byte, len(c.Modules))
	for _, module := range c.Modules {
		modulePath, err := confinedPath(caseDir, module.Path)
		if err != nil {
			return Case{}, nil, fmt.Errorf("oracle case %q module %q: %w", c.ID, module.Name, err)
		}
		source, err := os.ReadFile(modulePath)
		if err != nil {
			return Case{}, nil, fmt.Errorf("read oracle module %q: %w", module.Path, err)
		}
		matches := vbNamePattern.FindSubmatch(bytes.TrimPrefix(source, []byte{0xEF, 0xBB, 0xBF}))
		if len(matches) < 2 || !strings.EqualFold(string(matches[1]), module.Name) {
			return Case{}, nil, fmt.Errorf("oracle module %q Attribute VB_Name does not match declared name %q", module.Path, module.Name)
		}
		sources[module.Name] = source
	}
	return c, sources, nil
}

func ValidateCase(c Case, manifestID, caseDir string) error {
	if c.SchemaVersion != SchemaVersion {
		return fmt.Errorf("oracle case %q has unsupported schema_version %d", manifestID, c.SchemaVersion)
	}
	if c.ID != manifestID || !caseIDPattern.MatchString(c.ID) {
		return fmt.Errorf("oracle case id %q does not match manifest entry %q", c.ID, manifestID)
	}
	if c.Probe.Mode != ProbeCompile {
		return fmt.Errorf("oracle case %q: probe.mode must be %q", c.ID, ProbeCompile)
	}
	if len(c.Modules) == 0 {
		return fmt.Errorf("oracle case %q has no modules", c.ID)
	}
	seenNames := map[string]struct{}{}
	for _, module := range c.Modules {
		if strings.ToLower(strings.TrimSpace(module.Kind)) != "standard" {
			return fmt.Errorf("oracle case %q module %q: only standard modules are supported", c.ID, module.Name)
		}
		if module.Name == "" || module.Path == "" || strings.ToLower(filepath.Ext(module.Path)) != ".bas" {
			return fmt.Errorf("oracle case %q module has invalid name/path", c.ID)
		}
		key := strings.ToLower(module.Name)
		if _, ok := seenNames[key]; ok {
			return fmt.Errorf("oracle case %q has duplicate module name %q", c.ID, module.Name)
		}
		seenNames[key] = struct{}{}
		if _, err := confinedPath(caseDir, module.Path); err != nil {
			return fmt.Errorf("oracle case %q module %q: %w", c.ID, module.Name, err)
		}
	}
	if err := validateDiagnosticExpectations(c.Analysis.ExpectedDiagnostics, "expected", c.ID); err != nil {
		return err
	}
	if err := validateDiagnosticExpectations(c.Analysis.ForbiddenDiagnostics, "forbidden", c.ID); err != nil {
		return err
	}
	switch c.VBE.Expected {
	case ExpectedObserve:
		if strings.TrimSpace(c.Provenance.Status) != "pending" {
			return fmt.Errorf("oracle case %q: observe fixture provenance.status must be pending", c.ID)
		}
		if len(c.Provenance.VerifiedOn) != 0 {
			return fmt.Errorf("oracle case %q: pending fixture must not contain verified provenance", c.ID)
		}
		if c.VBE.EvidencePhase != EvidenceUnknown {
			return fmt.Errorf("oracle case %q: observe fixture evidence_phase must be unknown", c.ID)
		}
		if c.VBE.DiagnosticMeaning != MeaningObservation {
			return fmt.Errorf("oracle case %q: observe fixture diagnostic_meaning must be observation", c.ID)
		}
	case ExpectedAccepted, ExpectedRejected:
		if c.VBE.EvidencePhase != EvidenceCompile {
			return fmt.Errorf("oracle case %q: asserted fixture evidence_phase must be compile", c.ID)
		}
		if strings.TrimSpace(c.Provenance.Status) != "asserted" || len(c.Provenance.VerifiedOn) == 0 {
			return fmt.Errorf("oracle case %q: asserted fixture requires verified provenance", c.ID)
		}
		for _, metadata := range c.Provenance.VerifiedOn {
			if strings.TrimSpace(metadata.ExcelVersion) == "" || strings.TrimSpace(metadata.ExcelBuild) == "" || strings.TrimSpace(metadata.Bitness) == "" || strings.TrimSpace(metadata.Locale) == "" {
				return fmt.Errorf("oracle case %q: asserted provenance metadata is incomplete", c.ID)
			}
			verifiedAt, err := time.Parse(time.RFC3339, metadata.VerifiedAt)
			if err != nil {
				return fmt.Errorf("oracle case %q: asserted provenance verified_at must be RFC3339", c.ID)
			}
			if _, offset := verifiedAt.Zone(); offset != 0 {
				return fmt.Errorf("oracle case %q: asserted provenance verified_at must be UTC", c.ID)
			}
		}
		if c.VBE.Expected == ExpectedRejected && c.VBE.DiagnosticMeaning != MeaningCompileError {
			return fmt.Errorf("oracle case %q: rejected fixture diagnostic_meaning must be compile-error", c.ID)
		}
		if c.VBE.Expected == ExpectedAccepted && c.VBE.DiagnosticMeaning != MeaningSpecification && c.VBE.DiagnosticMeaning != MeaningPolicy && c.VBE.DiagnosticMeaning != MeaningMaintainability {
			return fmt.Errorf("oracle case %q: accepted fixture has invalid diagnostic_meaning", c.ID)
		}
	default:
		return fmt.Errorf("oracle case %q has invalid vbe.expected %q", c.ID, c.VBE.Expected)
	}
	return nil
}

func validateDiagnosticExpectations(expectations []DiagnosticExpectation, kind, caseID string) error {
	for _, expectation := range expectations {
		if strings.TrimSpace(expectation.Code) == "" {
			return fmt.Errorf("oracle case %q has an empty %s diagnostic code", caseID, kind)
		}
		if severity := strings.TrimSpace(expectation.Severity); severity != "" && severity != "error" && severity != "warning" {
			return fmt.Errorf("oracle case %q diagnostic %q has invalid severity %q", caseID, expectation.Code, expectation.Severity)
		}
		seenSurfaces := map[string]struct{}{}
		for _, surface := range expectation.Surfaces {
			surface = strings.TrimSpace(surface)
			if surface != "lint" && surface != "analyze" && surface != "lsp" {
				return fmt.Errorf("oracle case %q diagnostic %q has invalid surface %q", caseID, expectation.Code, surface)
			}
			if _, duplicate := seenSurfaces[surface]; duplicate {
				return fmt.Errorf("oracle case %q diagnostic %q repeats surface %q", caseID, expectation.Code, surface)
			}
			seenSurfaces[surface] = struct{}{}
		}
		if expectation.Range != nil {
			if expectation.Range.StartLine < 0 || expectation.Range.StartColumn < 0 || expectation.Range.EndLine < 0 || expectation.Range.EndColumn < 0 {
				return fmt.Errorf("oracle case %q diagnostic %q has a negative source range", caseID, expectation.Code)
			}
		}
	}
	return nil
}

func confinedPath(root, relative string) (string, error) {
	if filepath.IsAbs(relative) {
		return "", errors.New("path must be relative")
	}
	root, err := filepath.Abs(root)
	if err != nil {
		return "", err
	}
	path, err := filepath.Abs(filepath.Join(root, filepath.FromSlash(relative)))
	if err != nil {
		return "", err
	}
	rel, err := filepath.Rel(root, path)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return "", errors.New("path escapes fixture directory")
	}
	// Reject an existing symlink that resolves outside the fixture root while
	// retaining lexical containment checks for paths that will be created later.
	if resolvedRoot, resolveErr := filepath.EvalSymlinks(root); resolveErr == nil {
		if resolvedPath, resolveErr := filepath.EvalSymlinks(path); resolveErr == nil {
			resolvedRel, relErr := filepath.Rel(resolvedRoot, resolvedPath)
			if relErr != nil || resolvedRel == ".." || strings.HasPrefix(resolvedRel, ".."+string(filepath.Separator)) {
				return "", errors.New("path escapes fixture directory through symlink")
			}
		}
	}
	return path, nil
}

// SortedCaseIDs is useful for deterministic diagnostics and tests.
func SortedCaseIDs(cases []Case) []string {
	ids := make([]string, 0, len(cases))
	for _, c := range cases {
		ids = append(ids, c.ID)
	}
	sort.Strings(ids)
	return ids
}
