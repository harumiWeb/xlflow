package typedb

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"

	"github.com/harumiWeb/xlflow/internal/vbadb"
)

const (
	GeneratorName = "xlflow"
	EnvDir        = "XLFLOW_TYPE_DB_DIR"
)

type Manifest struct {
	SchemaVersion    int               `json:"schema_version"`
	Generator        string            `json:"generator"`
	GeneratorVersion string            `json:"generator_version"`
	GeneratedAt      string            `json:"generated_at"`
	Platform         string            `json:"platform"`
	Arch             string            `json:"arch"`
	Libraries        []ManifestLibrary `json:"libraries"`
}

type ManifestLibrary struct {
	Name   string `json:"name"`
	LibID  string `json:"libid"`
	Major  int    `json:"major"`
	Minor  int    `json:"minor"`
	LCID   int    `json:"lcid"`
	Source string `json:"source"`
	Output string `json:"output"`
}

type LibraryStatus struct {
	ManifestLibrary
	OutputPath string `json:"output_path"`
	Exists     bool   `json:"exists"`
	Stale      bool   `json:"stale"`
	Reason     string `json:"reason,omitempty"`
}

type Status struct {
	SchemaVersion    int             `json:"schema_version"`
	Generator        string          `json:"generator"`
	GeneratorVersion string          `json:"generator_version"`
	Dir              string          `json:"dir"`
	ManifestPath     string          `json:"manifest_path"`
	ManifestExists   bool            `json:"manifest_exists"`
	GeneratedAt      string          `json:"generated_at,omitempty"`
	Platform         string          `json:"platform"`
	Arch             string          `json:"arch"`
	GeneratedFiles   []string        `json:"generated_files"`
	Libraries        []LibraryStatus `json:"libraries"`
	Stale            bool            `json:"stale"`
	Reason           string          `json:"reason,omitempty"`
	SearchOrder      []string        `json:"search_order"`
}

type Options struct {
	Dir              string
	GeneratorVersion string
}

type LoadResult struct {
	DB             *vbadb.DB
	GeneratedDir   string
	GeneratedFiles []string
	Generated      bool
	Warnings       []string
}

func DefaultDir() (string, error) {
	if dir := strings.TrimSpace(os.Getenv(EnvDir)); dir != "" {
		return filepath.Clean(dir), nil
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(home, ".xlflow", "typelib"), nil
}

func ResolveDir(dir string) (string, error) {
	if strings.TrimSpace(dir) != "" {
		return filepath.Clean(dir), nil
	}
	return DefaultDir()
}

func ManifestPath(dir string) string {
	return filepath.Join(dir, "manifest.json")
}

func ReadManifest(dir string) (Manifest, error) {
	body, err := os.ReadFile(ManifestPath(dir))
	if err != nil {
		return Manifest{}, err
	}
	var manifest Manifest
	if err := json.Unmarshal(body, &manifest); err != nil {
		return Manifest{}, err
	}
	return manifest, nil
}

func WriteManifest(dir string, manifest Manifest) error {
	manifest.SchemaVersion = vbadb.SchemaVersion
	if manifest.Generator == "" {
		manifest.Generator = GeneratorName
	}
	if manifest.GeneratedAt == "" {
		manifest.GeneratedAt = time.Now().UTC().Format(time.RFC3339)
	}
	if manifest.Platform == "" {
		manifest.Platform = runtime.GOOS
	}
	if manifest.Arch == "" {
		manifest.Arch = runtime.GOARCH
	}
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return err
	}
	body, err := json.MarshalIndent(manifest, "", "  ")
	if err != nil {
		return err
	}
	body = append(body, '\n')
	return os.WriteFile(ManifestPath(dir), body, 0o644)
}

func StatusFor(opts Options) (Status, error) {
	dir, err := ResolveDir(opts.Dir)
	if err != nil {
		return Status{}, err
	}
	status := Status{
		SchemaVersion:    vbadb.SchemaVersion,
		Generator:        GeneratorName,
		GeneratorVersion: opts.GeneratorVersion,
		Dir:              dir,
		ManifestPath:     ManifestPath(dir),
		Platform:         runtime.GOOS,
		Arch:             runtime.GOARCH,
		SearchOrder: []string{
			"embedded built-in core DB",
			"global generated TypeLib DB",
			"xlflow curated overlay",
			"project-local type DB / user override",
		},
	}
	status.GeneratedFiles = generatedFiles(dir)
	manifest, err := ReadManifest(dir)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			status.Stale = true
			status.Reason = "manifest_missing"
			return status, nil
		}
		return Status{}, err
	}
	status.ManifestExists = true
	status.GeneratedAt = manifest.GeneratedAt
	status.Platform = manifest.Platform
	status.Arch = manifest.Arch
	status.Generator = firstNonEmpty(manifest.Generator, status.Generator)
	status.GeneratorVersion = firstNonEmpty(manifest.GeneratorVersion, status.GeneratorVersion)
	if manifest.SchemaVersion != vbadb.SchemaVersion {
		status.Stale = true
		status.Reason = "schema_version_changed"
	}
	if opts.GeneratorVersion != "" && manifest.GeneratorVersion != "" && manifest.GeneratorVersion != opts.GeneratorVersion {
		status.Stale = true
		if status.Reason == "" {
			status.Reason = "generator_version_changed"
		}
	}
	for _, library := range manifest.Libraries {
		libStatus := LibraryStatus{
			ManifestLibrary: library,
			OutputPath:      filepath.Join(dir, library.Output),
		}
		if strings.TrimSpace(library.Output) == "" {
			libStatus.Stale = true
			libStatus.Reason = "output_missing"
		} else if _, err := os.Stat(libStatus.OutputPath); err == nil {
			libStatus.Exists = true
		} else if errors.Is(err, os.ErrNotExist) {
			libStatus.Stale = true
			libStatus.Reason = "output_file_missing"
		} else {
			return Status{}, err
		}
		if libStatus.Stale {
			status.Stale = true
			if status.Reason == "" {
				status.Reason = libStatus.Reason
			}
		}
		status.Libraries = append(status.Libraries, libStatus)
	}
	if len(status.Libraries) == 0 {
		status.Stale = true
		if status.Reason == "" {
			status.Reason = "no_libraries"
		}
	}
	return status, nil
}

func Clean(dir string) (string, error) {
	resolved, err := ResolveDir(dir)
	if err != nil {
		return "", err
	}
	if strings.TrimSpace(resolved) == "" || filepath.Clean(resolved) == filepath.Clean(string(os.PathSeparator)) {
		return "", fmt.Errorf("refusing to clean unsafe type DB directory %q", resolved)
	}
	if err := os.RemoveAll(resolved); err != nil {
		return "", err
	}
	return resolved, nil
}

func LoadGenerated(dir string) (*vbadb.DB, error) {
	resolved, err := ResolveDir(dir)
	if err != nil {
		return nil, err
	}
	if _, err := os.Stat(resolved); err != nil {
		return vbadb.New(), nil
	}
	return vbadb.LoadDir(resolved)
}

func LoadForRuntime(dir string) (LoadResult, error) {
	resolved, err := ResolveDir(dir)
	if err != nil {
		return LoadResult{}, err
	}
	result := LoadResult{GeneratedDir: resolved}
	files := generatedFiles(resolved)
	result.GeneratedFiles = files
	db := vbadb.New()
	if len(files) > 0 {
		generated, err := vbadb.LoadFiles(files...)
		if err != nil {
			result.Warnings = append(result.Warnings, "generated type database could not be loaded: "+err.Error())
		} else {
			db = generated
			result.Generated = true
		}
	}
	if err := vbadb.LoadBuiltinInto(db); err != nil {
		return LoadResult{}, err
	}
	result.DB = db
	return result, nil
}

func generatedFiles(dir string) []string {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil
	}
	var files []string
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if strings.EqualFold(name, "manifest.json") || !strings.EqualFold(filepath.Ext(name), ".json") {
			continue
		}
		files = append(files, filepath.Join(dir, name))
	}
	sort.Strings(files)
	return files
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}
