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

	staticcontract "github.com/harumiWeb/xlflow/internal/staticanalysis/contract"
	"github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
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

	BindingUnbound        = "unbound"
	BindingPartiallyBound = "partially-bound"
	BindingBound          = "bound"
	BindingNotApplicable  = "not-applicable"

	EvidenceRoleCompileEquivalent          = "compile-equivalent"
	EvidenceRolePolicyObservation          = "policy-observation"
	EvidenceRoleMaintainabilityObservation = "maintainability-observation"
	EvidenceRoleLanguageObservation        = "language-observation"
	EvidenceRoleHarnessControl             = "harness-control"
)

var caseIDPattern = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._-]*$`)
var vbNamePattern = regexp.MustCompile(`(?im)^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"`)
var libIDPattern = regexp.MustCompile(`^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$`)

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
	References    []Reference         `json:"references,omitempty"`
	Modules       []Module            `json:"modules"`
	Probe         Probe               `json:"probe"`
	VBE           VBEExpectation      `json:"vbe"`
	Analysis      AnalysisExpectation `json:"analysis,omitempty"`
	Provenance    Provenance          `json:"provenance"`
}

// Reference identifies a registered VBA project TypeLib to add before the
// fixture modules are compiled. Name is descriptive; the LIBID and version
// are the values passed to VBProject.References.AddFromGuid.
type Reference struct {
	Name  string `json:"name,omitempty"`
	LibID string `json:"libid"`
	Major int    `json:"major"`
	Minor int    `json:"minor"`
}

type Module struct {
	Name  string `json:"name"`
	Kind  string `json:"kind"`
	Path  string `json:"path"`
	Entry bool   `json:"entry,omitempty"`
	// DocumentTarget identifies the host for a document module.  It is an
	// additive v1 field: standard-module fixtures can omit it, while document
	// modules use "workbook" or "worksheet" to make the VBIDE host explicit.
	DocumentTarget string `json:"document_target,omitempty"`
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
	BindingStatus        string                  `json:"binding_status"`
	EvidenceRole         string                  `json:"evidence_role"`
	RuleCodes            []string                `json:"rule_codes,omitempty"`
	BindingNote          *string                 `json:"binding_note,omitempty"`
	NegativeControls     []string                `json:"negative_controls,omitempty"`
	ExpectedDiagnostics  []DiagnosticExpectation `json:"expected_diagnostics,omitempty"`
	ForbiddenDiagnostics []DiagnosticExpectation `json:"forbidden_diagnostics,omitempty"`
}

type DiagnosticExpectation = staticcontract.DiagnosticExpectation

type Range = staticcontract.Range

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
	if err := validateReferences(c.References, c.ID); err != nil {
		return err
	}
	seenNames := map[string]struct{}{}
	for _, module := range c.Modules {
		kind := strings.ToLower(strings.TrimSpace(module.Kind))
		if kind != "standard" && kind != "class" && kind != "form" && kind != "document" {
			return fmt.Errorf("oracle case %q module %q: unsupported module kind %q", c.ID, module.Name, module.Kind)
		}
		if module.Name == "" || module.Path == "" {
			return fmt.Errorf("oracle case %q module has invalid name/path", c.ID)
		}
		ext := strings.ToLower(filepath.Ext(module.Path))
		// Oracle fixtures deliberately keep every module kind in ordinary .bas
		// files. The bridge provisions a component and injects its sanitized
		// body, so designer .frm/.frx artifacts are outside this contract.
		validExt := ext == ".bas"
		if !validExt {
			return fmt.Errorf("oracle case %q module %q: kind %q requires a compatible source extension", c.ID, module.Name, kind)
		}
		target := strings.ToLower(strings.TrimSpace(module.DocumentTarget))
		if kind == "document" {
			if target != "workbook" && target != "worksheet" {
				return fmt.Errorf("oracle case %q document module %q: document_target must be workbook or worksheet", c.ID, module.Name)
			}
		} else if target != "" {
			return fmt.Errorf("oracle case %q module %q: document_target is only valid for document modules", c.ID, module.Name)
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
	if err := validateAnalysisBinding(c); err != nil {
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

func validateReferences(references []Reference, caseID string) error {
	seen := map[string]struct{}{}
	for _, reference := range references {
		libID := strings.TrimSpace(reference.LibID)
		if strings.HasPrefix(libID, "{") != strings.HasSuffix(libID, "}") {
			return fmt.Errorf("oracle case %q reference %q has an invalid libid", caseID, reference.Name)
		}
		libID = strings.TrimPrefix(libID, "{")
		libID = strings.TrimSuffix(libID, "}")
		if !libIDPattern.MatchString(libID) {
			return fmt.Errorf("oracle case %q reference %q has an invalid libid", caseID, reference.Name)
		}
		if reference.Major < 0 || reference.Minor < 0 {
			return fmt.Errorf("oracle case %q reference %q has a negative version", caseID, reference.Name)
		}
		key := strings.ToLower(libID) + fmt.Sprintf("/%d/%d", reference.Major, reference.Minor)
		if _, ok := seen[key]; ok {
			return fmt.Errorf("oracle case %q repeats reference %q version %d.%d", caseID, reference.Name, reference.Major, reference.Minor)
		}
		seen[key] = struct{}{}
	}
	return nil
}

func validateAnalysisBinding(c Case) error {
	analysis := c.Analysis
	switch analysis.BindingStatus {
	case BindingUnbound, BindingPartiallyBound, BindingBound, BindingNotApplicable:
	default:
		return fmt.Errorf("oracle case %q has invalid analysis.binding_status %q", c.ID, analysis.BindingStatus)
	}
	if err := validateEvidenceRole(c); err != nil {
		return err
	}

	seenCodes := make(map[string]struct{}, len(analysis.RuleCodes))
	for _, code := range analysis.RuleCodes {
		if strings.TrimSpace(code) == "" || code != strings.TrimSpace(code) {
			return fmt.Errorf("oracle case %q has an empty or padded analysis rule code", c.ID)
		}
		rule, err := staticcontract.CanonicalRuleMetadata(code)
		if err != nil {
			return fmt.Errorf("oracle case %q has invalid analysis rule code %q: %w", c.ID, code, err)
		}
		if analysis.EvidenceRole == EvidenceRoleCompileEquivalent && !rule.CompileEquivalent {
			// An accepted, fully-bound fixture may also carry the explicit
			// non-blocking style projection for the same source. It remains a
			// supplemental policy/style contract, not compiler evidence.
			if !supplementalAcceptedStyleBinding(c, code) {
				return fmt.Errorf("oracle case %q binds non-compile-equivalent rule %q as compile-equivalent", c.ID, code)
			}
		}
		if _, ok := seenCodes[code]; ok {
			return fmt.Errorf("oracle case %q repeats analysis rule code %q", c.ID, code)
		}
		seenCodes[code] = struct{}{}
	}
	if analysis.BindingNote != nil && strings.TrimSpace(*analysis.BindingNote) == "" {
		return fmt.Errorf("oracle case %q has an empty analysis.binding_note", c.ID)
	}

	hasExpected := len(analysis.ExpectedDiagnostics) > 0
	hasForbidden := len(analysis.ForbiddenDiagnostics) > 0
	switch analysis.BindingStatus {
	case BindingUnbound:
		if hasExpected || hasForbidden {
			return fmt.Errorf("oracle case %q: unbound fixture cannot declare analyzer diagnostic contracts", c.ID)
		}
		if len(analysis.RuleCodes) > 0 && analysis.BindingNote == nil {
			return fmt.Errorf("oracle case %q: unbound fixture with rule codes requires analysis.binding_note", c.ID)
		}
	case BindingPartiallyBound:
		if len(analysis.RuleCodes) == 0 {
			return fmt.Errorf("oracle case %q: partially-bound fixture requires analysis.rule_codes", c.ID)
		}
		if analysis.BindingNote == nil {
			return fmt.Errorf("oracle case %q: partially-bound fixture requires analysis.binding_note", c.ID)
		}
	case BindingBound:
		if len(analysis.RuleCodes) == 0 {
			return fmt.Errorf("oracle case %q: bound fixture requires analysis.rule_codes", c.ID)
		}
		var boundExpectations []DiagnosticExpectation
		switch c.VBE.Expected {
		case ExpectedRejected:
			if !hasExpected {
				return fmt.Errorf("oracle case %q: bound rejected fixture requires expected diagnostics", c.ID)
			}
			boundExpectations = analysis.ExpectedDiagnostics
		case ExpectedAccepted:
			if !hasForbidden {
				return fmt.Errorf("oracle case %q: bound accepted fixture requires forbidden diagnostics", c.ID)
			}
			boundExpectations = analysis.ForbiddenDiagnostics
		default:
			return fmt.Errorf("oracle case %q: bound fixture requires asserted VBE evidence", c.ID)
		}
		for _, code := range analysis.RuleCodes {
			if !diagnosticCodeDeclared(code, boundExpectations) {
				if supplementalAcceptedStyleBinding(c, code) && diagnosticCodeDeclared(code, analysis.ExpectedDiagnostics) {
					continue
				}
				return fmt.Errorf("oracle case %q: bound rule code %q is not declared by an analysis contract", c.ID, code)
			}
			if c.VBE.Expected == ExpectedRejected {
				for _, expectation := range analysis.ExpectedDiagnostics {
					if expectation.Code == code && expectation.Severity != string(rules.SeverityError) {
						return fmt.Errorf("oracle case %q: rejected compile-equivalent rule %q must expect severity %q", c.ID, code, rules.SeverityError)
					}
				}
			}
		}
	case BindingNotApplicable:
		if len(analysis.RuleCodes) > 0 || hasExpected || hasForbidden {
			return fmt.Errorf("oracle case %q: not-applicable fixture cannot declare analyzer bindings", c.ID)
		}
	}
	if analysis.BindingStatus == BindingBound || analysis.BindingStatus == BindingPartiallyBound {
		contracts := append(append([]DiagnosticExpectation(nil), analysis.ExpectedDiagnostics...), analysis.ForbiddenDiagnostics...)
		if analysis.BindingStatus == BindingBound {
			for _, code := range analysis.RuleCodes {
				if !diagnosticCodeDeclared(code, contracts) {
					return fmt.Errorf("oracle case %q: analysis rule code %q is not declared by an expected or forbidden diagnostic contract", c.ID, code)
				}
			}
		}
		for _, expectation := range contracts {
			if _, ok := seenCodes[expectation.Code]; !ok {
				if analysis.BindingStatus == BindingBound {
					return fmt.Errorf("oracle case %q: bound diagnostic code %q is not declared by analysis.rule_codes", c.ID, expectation.Code)
				}
				return fmt.Errorf("oracle case %q: partially-bound diagnostic code %q is not declared by analysis.rule_codes", c.ID, expectation.Code)
			}
		}
	}
	return nil
}

func validateEvidenceRole(c Case) error {
	role := strings.TrimSpace(c.Analysis.EvidenceRole)
	if role != c.Analysis.EvidenceRole {
		return fmt.Errorf("oracle case %q has padded analysis.evidence_role %q", c.ID, c.Analysis.EvidenceRole)
	}
	switch role {
	case EvidenceRoleCompileEquivalent, EvidenceRolePolicyObservation,
		EvidenceRoleMaintainabilityObservation, EvidenceRoleLanguageObservation, EvidenceRoleHarnessControl:
	case "":
		return fmt.Errorf("oracle case %q: analysis.evidence_role is required", c.ID)
	default:
		return fmt.Errorf("oracle case %q has invalid analysis.evidence_role %q", c.ID, c.Analysis.EvidenceRole)
	}

	switch c.Analysis.BindingStatus {
	case BindingPartiallyBound:
		if role != EvidenceRoleCompileEquivalent {
			return fmt.Errorf("oracle case %q: partially-bound fixture requires evidence_role %q", c.ID, EvidenceRoleCompileEquivalent)
		}
	case BindingBound:
		if role != EvidenceRoleCompileEquivalent {
			return fmt.Errorf("oracle case %q: bound fixture requires evidence_role %q", c.ID, EvidenceRoleCompileEquivalent)
		}
	case BindingNotApplicable:
		if role != EvidenceRoleHarnessControl {
			return fmt.Errorf("oracle case %q: not-applicable fixture requires evidence_role %q", c.ID, EvidenceRoleHarnessControl)
		}
	case BindingUnbound:
		if role == EvidenceRoleHarnessControl {
			return fmt.Errorf("oracle case %q: unbound fixture cannot use evidence_role %q", c.ID, EvidenceRoleHarnessControl)
		}
	}
	return nil
}

func diagnosticCodeDeclared(code string, expectations []DiagnosticExpectation) bool {
	for _, expectation := range expectations {
		if expectation.Code == code {
			return true
		}
	}
	return false
}

func validateDiagnosticExpectations(expectations []DiagnosticExpectation, kind, caseID string) error {
	if err := staticcontract.ValidateExpectations(expectations); err != nil {
		return fmt.Errorf("oracle case %q has invalid %s diagnostic contract: %w", caseID, kind, err)
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
