// Package corpus owns the pinned, third-party static-analysis corpus manifest
// and the developer-only synchronization operation used to materialize it.
//
// The package deliberately has no production CLI dependencies.  Callers that
// want to update the vendored tree must invoke Sync explicitly; loading and
// validating a manifest never accesses the network.
package corpus

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
)

const SchemaVersion = 2

const (
	ProfileExcel        = "excel"
	ProfileGenericVBA   = "generic-vba"
	ProfileAccess       = "access"
	OriginTreeSitterVBA = "tree-sitter-vba"
)

// Manifest is the schema-versioned source of truth for the corpus.
type Manifest struct {
	SchemaVersion int       `json:"schema_version"`
	Upstream      Upstream  `json:"upstream"`
	Projects      []Project `json:"projects"`
	jsonLoaded    bool
}

type Upstream struct {
	Repository string `json:"repository"`
	Commit     string `json:"commit"`
}

type Project struct {
	ID              string           `json:"id"`
	Path            string           `json:"path"`
	Profile         string           `json:"profile"`
	Enabled         bool             `json:"enabled"`
	Notes           string           `json:"notes,omitempty"`
	Source          Source           `json:"source"`
	SourceCounts    SourceCounts     `json:"source_counts"`
	Provenance      Provenance       `json:"provenance"`
	Classifications []Classification `json:"classifications,omitempty"`

	// Presence flags are populated while decoding JSON. They make it possible to
	// distinguish deliberately supplied zero/false values from missing fields,
	// while retaining convenient value fields for callers constructing values.
	enabledPresent      bool
	notesPresent        bool
	sourceCountsPresent bool
}

// SourceCounts records the expected number of exported VBA source files in a
// project. Counts are part of the manifest contract so a refresh cannot
// silently drop or add source files while preserving the same project ID.
type SourceCounts struct {
	Bas int `json:"bas"`
	Cls int `json:"cls"`
	Frm int `json:"frm"`
}

// Total returns the number of VBA source files represented by the counts.
func (c SourceCounts) Total() int { return c.Bas + c.Cls + c.Frm }

// Classification overrides the default extension-to-module-kind mapping for
// one source file. Paths are project-relative and deliberately exact; the
// adapter never guesses document-module semantics from file names or VBA
// attributes.
type Classification struct {
	Path string `json:"path"`
	Kind string `json:"kind"`
}

const (
	ModuleKindStandard = "standard"
	ModuleKindClass    = "class"
	ModuleKindForm     = "form"
	ModuleKindDocument = "document"
)

type Source struct {
	Origin string `json:"origin"`
	Path   string `json:"path"`
}

type Provenance struct {
	Repository  string `json:"repository"`
	License     string `json:"license"`
	LicenseFile string `json:"license_file"`
	SourceFile  string `json:"source_file"`
}

func (m *Manifest) UnmarshalJSON(data []byte) error {
	type alias Manifest
	var value alias
	if err := decodeStrict(data, &value); err != nil {
		return err
	}
	value.jsonLoaded = true
	*m = Manifest(value)
	return nil
}

// UnmarshalJSON rejects unknown fields and records presence of optional JSON
// values needed by the manifest contract.
func (p *Project) UnmarshalJSON(data []byte) error {
	type alias Project
	var value alias
	if err := decodeStrict(data, &value); err != nil {
		return err
	}
	var fields map[string]json.RawMessage
	if err := json.Unmarshal(data, &fields); err != nil {
		return err
	}
	enabled, ok := fields["enabled"]
	value.enabledPresent = ok
	if ok {
		var parsed bool
		if string(enabled) == "null" || json.Unmarshal(enabled, &parsed) != nil {
			return errors.New("project enabled must be a boolean")
		}
	}
	_, value.notesPresent = fields["notes"]
	counts, ok := fields["source_counts"]
	value.sourceCountsPresent = ok
	if ok {
		if string(counts) == "null" {
			return errors.New("project source_counts must be an object")
		}
		var countFields map[string]json.RawMessage
		if err := json.Unmarshal(counts, &countFields); err != nil || countFields == nil {
			if err != nil {
				return fmt.Errorf("project source_counts must be an object: %w", err)
			}
			return errors.New("project source_counts must be an object")
		}
		for _, name := range []string{"bas", "cls", "frm"} {
			raw, present := countFields[name]
			if !present {
				return fmt.Errorf("project source_counts must declare %s", name)
			}
			if string(raw) == "null" {
				return fmt.Errorf("project source_counts.%s must be an integer", name)
			}
		}
		var parsed SourceCounts
		if err := decodeStrict(counts, &parsed); err != nil {
			return fmt.Errorf("project source_counts must contain integer bas/cls/frm counts: %w", err)
		}
		value.SourceCounts = parsed
	}
	*p = Project(value)
	return nil
}

var (
	commitPattern     = regexp.MustCompile(`^[0-9a-f]{40}$`)
	repositoryPattern = regexp.MustCompile(`^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$`)
	identifierPattern = regexp.MustCompile(`^[a-z0-9][a-z0-9._-]*$`)
)

// LoadManifest reads and strictly validates a manifest.  The returned root is
// the directory containing the manifest and is useful to resolve its corpus.
func LoadManifest(manifestPath string) (Manifest, string, error) {
	abs, err := filepath.Abs(manifestPath)
	if err != nil {
		return Manifest{}, "", fmt.Errorf("resolve corpus manifest: %w", err)
	}
	body, err := os.ReadFile(abs)
	if err != nil {
		return Manifest{}, "", fmt.Errorf("read corpus manifest: %w", err)
	}
	var manifest Manifest
	if err := decodeStrict(body, &manifest); err != nil {
		return Manifest{}, "", fmt.Errorf("decode corpus manifest: %w", err)
	}
	if err := ValidateManifest(manifest); err != nil {
		return Manifest{}, "", err
	}
	return manifest, filepath.Dir(abs), nil
}

// Validate is an alias for ValidateManifest for callers that use the shorter
// validation convention used by the other static-analysis registries.
func Validate(manifest Manifest) error { return ValidateManifest(manifest) }

func ValidateManifest(manifest Manifest) error {
	if manifest.SchemaVersion != SchemaVersion {
		return fmt.Errorf("unsupported corpus manifest schema_version %d (want %d)", manifest.SchemaVersion, SchemaVersion)
	}
	if !repositoryPattern.MatchString(manifest.Upstream.Repository) || strings.Contains(manifest.Upstream.Repository, "..") {
		return fmt.Errorf("upstream.repository must be a GitHub owner/repository, got %q", manifest.Upstream.Repository)
	}
	if !commitPattern.MatchString(manifest.Upstream.Commit) {
		return fmt.Errorf("upstream.commit must be a lowercase 40-character SHA, got %q", manifest.Upstream.Commit)
	}
	if len(manifest.Projects) == 0 {
		return errors.New("corpus manifest contains no projects")
	}
	seenIDs := make(map[string]struct{}, len(manifest.Projects))
	seenDest := make(map[string]string, len(manifest.Projects))
	seenSource := make(map[string]string, len(manifest.Projects))
	for i := range manifest.Projects {
		p := manifest.Projects[i]
		if !identifierPattern.MatchString(p.ID) || p.ID != strings.TrimSpace(p.ID) {
			return fmt.Errorf("project %d has invalid id %q", i, p.ID)
		}
		if _, ok := seenIDs[p.ID]; ok {
			return fmt.Errorf("duplicate project id %q", p.ID)
		}
		seenIDs[p.ID] = struct{}{}
		if i > 0 && manifest.Projects[i-1].ID >= p.ID {
			return fmt.Errorf("projects must be sorted by id: %q follows %q", p.ID, manifest.Projects[i-1].ID)
		}
		if manifest.jsonLoaded && !p.enabledPresent {
			return fmt.Errorf("project %q must explicitly declare enabled", p.ID)
		}
		if manifest.jsonLoaded && !p.sourceCountsPresent {
			return fmt.Errorf("project %q must explicitly declare source_counts", p.ID)
		}
		if p.SourceCounts.Bas < 0 || p.SourceCounts.Cls < 0 || p.SourceCounts.Frm < 0 {
			return fmt.Errorf("project %q source_counts must be non-negative", p.ID)
		}
		if err := validateRelativePath(p.Path); err != nil {
			return fmt.Errorf("project %q destination path: %w", p.ID, err)
		}
		if !strings.HasPrefix(p.Path, "projects/third_party/") || p.Path == "projects/third_party/" {
			return fmt.Errorf("project %q destination path must be below projects/third_party", p.ID)
		}
		destKey := strings.ToLower(p.Path)
		if prior, ok := seenDest[destKey]; ok {
			return fmt.Errorf("projects %q and %q share destination path %q", prior, p.ID, p.Path)
		}
		seenDest[destKey] = p.ID
		if p.Profile != ProfileExcel && p.Profile != ProfileGenericVBA && p.Profile != ProfileAccess {
			return fmt.Errorf("project %q has unsupported profile %q", p.ID, p.Profile)
		}
		if p.notesPresent && (p.Notes == "" || p.Notes != strings.TrimSpace(p.Notes)) {
			return fmt.Errorf("project %q notes must be non-empty and unpadded", p.ID)
		}
		if !p.Enabled && strings.TrimSpace(p.Notes) == "" {
			return fmt.Errorf("disabled project %q requires a non-empty notes reason", p.ID)
		}
		if err := validateSourcePath(p.Source.Path); err != nil {
			return fmt.Errorf("project %q source.path: %w", p.ID, err)
		}
		if p.Source.Origin != OriginTreeSitterVBA {
			return fmt.Errorf("project %q has unsupported source.origin %q", p.ID, p.Source.Origin)
		}
		sourceKey := strings.ToLower(p.Source.Path)
		if prior, ok := seenSource[sourceKey]; ok {
			return fmt.Errorf("projects %q and %q share source path %q", prior, p.ID, p.Source.Path)
		}
		seenSource[sourceKey] = p.ID
		if err := validateProvenance(p.Provenance); err != nil {
			return fmt.Errorf("project %q provenance: %w", p.ID, err)
		}
		if err := validateClassifications(p.ID, p.Classifications); err != nil {
			return err
		}
	}
	// Source and destination paths may not overlap; otherwise one copy could
	// consume or overwrite another project and boundaries would be ambiguous.
	for i := range manifest.Projects {
		for j := i + 1; j < len(manifest.Projects); j++ {
			if pathOverlap(manifest.Projects[i].Source.Path, manifest.Projects[j].Source.Path) {
				return fmt.Errorf("projects %q and %q have overlapping source paths", manifest.Projects[i].ID, manifest.Projects[j].ID)
			}
			if pathOverlap(manifest.Projects[i].Path, manifest.Projects[j].Path) {
				return fmt.Errorf("projects %q and %q have overlapping destination paths", manifest.Projects[i].ID, manifest.Projects[j].ID)
			}
		}
	}
	return nil
}

func validateClassifications(projectID string, classifications []Classification) error {
	seen := make(map[string]struct{}, len(classifications))
	for i, classification := range classifications {
		if err := validateRelativePath(classification.Path); err != nil {
			return fmt.Errorf("project %q classification %d path: %w", projectID, i, err)
		}
		ext := strings.ToLower(path.Ext(classification.Path))
		switch ext {
		case ".bas", ".cls", ".frm":
		default:
			return fmt.Errorf("project %q classification %q targets unsupported extension %q", projectID, classification.Path, ext)
		}
		switch ext {
		case ".bas":
			if classification.Kind != ModuleKindStandard {
				return fmt.Errorf("project %q classification %q must use kind %q", projectID, classification.Path, ModuleKindStandard)
			}
		case ".frm":
			if classification.Kind != ModuleKindForm {
				return fmt.Errorf("project %q classification %q must use kind %q", projectID, classification.Path, ModuleKindForm)
			}
		case ".cls":
			if classification.Kind != ModuleKindClass && classification.Kind != ModuleKindDocument {
				return fmt.Errorf("project %q classification %q has unsupported kind %q", projectID, classification.Path, classification.Kind)
			}
		}
		key := strings.ToLower(classification.Path)
		if _, exists := seen[key]; exists {
			return fmt.Errorf("project %q has duplicate classification path %q", projectID, classification.Path)
		}
		seen[key] = struct{}{}
		if i > 0 && classifications[i-1].Path >= classification.Path {
			return fmt.Errorf("project %q classifications must be sorted by path: %q follows %q", projectID, classification.Path, classifications[i-1].Path)
		}
	}
	return nil
}

func validateProvenance(p Provenance) error {
	if strings.TrimSpace(p.Repository) == "" {
		return errors.New("repository is required")
	}
	u, err := url.ParseRequestURI(p.Repository)
	if err != nil || u.Scheme != "https" || u.Host == "" || u.User != nil {
		return fmt.Errorf("repository must be an HTTPS URL, got %q", p.Repository)
	}
	if strings.TrimSpace(p.License) == "" || strings.ContainsAny(p.License, " \t\r\n") {
		return fmt.Errorf("license must be a non-empty SPDX identifier, got %q", p.License)
	}
	if err := validateRelativePath(p.LicenseFile); err != nil {
		return fmt.Errorf("license_file: %w", err)
	}
	if err := validateRelativePath(p.SourceFile); err != nil {
		return fmt.Errorf("source_file: %w", err)
	}
	if strings.EqualFold(p.LicenseFile, p.SourceFile) {
		return errors.New("license_file and source_file must differ")
	}
	return nil
}

func validateSourcePath(value string) error {
	if err := validateRelativePath(value); err != nil {
		return err
	}
	if value == "examples/third_party" || !strings.HasPrefix(value, "examples/third_party/") {
		return fmt.Errorf("must be below examples/third_party")
	}
	return nil
}

func validateRelativePath(value string) error {
	if value == "" || strings.TrimSpace(value) != value || strings.Contains(value, "\\") || strings.Contains(value, ":") {
		return fmt.Errorf("must be a canonical slash-separated relative path, got %q", value)
	}
	if strings.HasPrefix(value, "/") || path.IsAbs(value) {
		return fmt.Errorf("must not be absolute")
	}
	clean := path.Clean(value)
	if clean != value || clean == "." || strings.HasPrefix(clean, "../") || clean == ".." {
		return fmt.Errorf("must not contain traversal or non-canonical segments")
	}
	for _, segment := range strings.Split(value, "/") {
		if segment == "" || segment == "." || segment == ".." {
			return fmt.Errorf("contains an invalid path segment")
		}
	}
	return nil
}

func pathOverlap(a, b string) bool {
	a, b = strings.ToLower(path.Clean(a)), strings.ToLower(path.Clean(b))
	return a == b || strings.HasPrefix(a, b+"/") || strings.HasPrefix(b, a+"/")
}

func decodeStrict(data []byte, out any) error {
	dec := json.NewDecoder(bytes.NewReader(data))
	dec.DisallowUnknownFields()
	if err := dec.Decode(out); err != nil {
		return err
	}
	var extra any
	if err := dec.Decode(&extra); err == nil {
		return errors.New("manifest contains trailing JSON")
	} else if !errors.Is(err, io.EOF) {
		return fmt.Errorf("manifest contains trailing invalid JSON: %w", err)
	}
	return nil
}

// CanonicalProjectIDs returns IDs in manifest order.  It is useful to callers
// displaying deterministic sync output and returns a defensive copy.
func CanonicalProjectIDs(manifest Manifest) []string {
	ids := make([]string, 0, len(manifest.Projects))
	for _, p := range manifest.Projects {
		ids = append(ids, p.ID)
	}
	return ids
}

// SortedProjectIDs is retained as a helper for tools that need to compare a
// manifest against an independently discovered project list.
func SortedProjectIDs(manifest Manifest) []string {
	ids := CanonicalProjectIDs(manifest)
	sort.Strings(ids)
	return ids
}
