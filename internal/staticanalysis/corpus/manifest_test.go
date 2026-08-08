package corpus

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestRepositoryManifestAndMetadata(t *testing.T) {
	manifest, root, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	if manifest.Upstream.Commit != "c867f27ea3dedc2ccece1eeb0273cdb242899182" {
		t.Fatalf("unexpected pinned commit %q", manifest.Upstream.Commit)
	}
	if got := CanonicalProjectIDs(manifest); strings.Join(got, ",") != "access-examples,vba-json,vba-web" {
		t.Fatalf("unexpected project order %v", got)
	}
	for _, project := range manifest.Projects {
		destination := filepath.Join(root, filepath.FromSlash(project.Path))
		for _, relative := range []string{project.Provenance.LicenseFile, project.Provenance.SourceFile} {
			if _, err := os.Stat(filepath.Join(destination, filepath.FromSlash(relative))); err != nil {
				t.Fatalf("project %s metadata %s: %v", project.ID, relative, err)
			}
		}
	}
}

func TestLoadManifestRejectsMalformedDocuments(t *testing.T) {
	valid := `{"schema_version":1,"upstream":{"repository":"harumiWeb/tree-sitter-vba","commit":"c867f27ea3dedc2ccece1eeb0273cdb242899182"},"projects":[{"id":"one","path":"projects/third_party/one","profile":"generic-vba","enabled":true,"source":{"origin":"tree-sitter-vba","path":"examples/third_party/one"},"provenance":{"repository":"https://example.test/one","license":"MIT","license_file":"LICENSE","source_file":"SOURCE.md"}}]}`
	tests := []struct {
		name string
		body string
		want string
	}{
		{"unknown field", strings.Replace(valid, `"enabled":true`, `"enabled":true,"unexpected":true`, 1), "unknown field"},
		{"unsupported schema", strings.Replace(valid, `"schema_version":1`, `"schema_version":2`, 1), "unsupported corpus manifest"},
		{"short commit", strings.Replace(valid, "c867f27ea3dedc2ccece1eeb0273cdb242899182", "c867f27", 1), "40-character SHA"},
		{"missing enabled", strings.Replace(valid, `,"enabled":true`, "", 1), "explicitly declare enabled"},
		{"null enabled", strings.Replace(valid, `"enabled":true`, `"enabled":null`, 1), "enabled must be a boolean"},
		{"path traversal", strings.Replace(valid, "projects/third_party/one", "projects/third_party/../one", 1), "non-canonical"},
		{"unsupported profile", strings.Replace(valid, `"generic-vba"`, `"other"`, 1), "unsupported profile"},
		{"empty notes", strings.Replace(valid, `"enabled":true`, `"enabled":true,"notes":"  "`, 1), "notes"},
		{"disabled without reason", strings.Replace(valid, `"enabled":true`, `"enabled":false`, 1), "requires a non-empty notes reason"},
		{"unsupported classification kind", strings.Replace(valid, `"source":{"origin"`, `"classifications":[{"path":"Main.cls","kind":"other"}],"source":{"origin"`, 1), "unsupported kind"},
		{"classification traversal", strings.Replace(valid, `"source":{"origin"`, `"classifications":[{"path":"../Main.bas","kind":"standard"}],"source":{"origin"`, 1), "classification"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			path := filepath.Join(t.TempDir(), "manifest.json")
			if err := os.WriteFile(path, []byte(test.body), 0o644); err != nil {
				t.Fatal(err)
			}
			_, _, err := LoadManifest(path)
			if err == nil || !strings.Contains(strings.ToLower(err.Error()), strings.ToLower(test.want)) {
				t.Fatalf("LoadManifest() error = %v, want substring %q", err, test.want)
			}
		})
	}
}

func TestValidateManifestRejectsInvalidClassificationOrderingAndDuplicates(t *testing.T) {
	manifest := fixtureManifest("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
	manifest.Projects[0].Classifications = []Classification{
		{Path: "z.cls", Kind: ModuleKindClass},
		{Path: "a.cls", Kind: ModuleKindClass},
	}
	if err := ValidateManifest(manifest); err == nil || !strings.Contains(err.Error(), "sorted by path") {
		t.Fatalf("unsorted classifications accepted: %v", err)
	}
	manifest.Projects[0].Classifications = []Classification{
		{Path: "Main.cls", Kind: ModuleKindClass},
		{Path: "main.cls", Kind: ModuleKindClass},
	}
	if err := ValidateManifest(manifest); err == nil || !strings.Contains(err.Error(), "duplicate classification") {
		t.Fatalf("case-insensitive duplicate classifications accepted: %v", err)
	}
}

func TestValidateManifestRejectsInvalidClassificationKindPairs(t *testing.T) {
	for _, test := range []struct {
		name, path, kind, want string
	}{
		{"standard source as document", "Main.bas", ModuleKindDocument, "must use kind \"standard\""},
		{"form source as class", "Dialog.frm", ModuleKindClass, "must use kind \"form\""},
		{"class source as form", "Sheet.cls", ModuleKindForm, "unsupported kind"},
	} {
		t.Run(test.name, func(t *testing.T) {
			manifest := fixtureManifest("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
			manifest.Projects[0].Classifications = []Classification{{Path: test.path, Kind: test.kind}}
			if err := ValidateManifest(manifest); err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("invalid classification pair error = %v, want %q", err, test.want)
			}
		})
	}
}

func TestValidateManifestRejectsCaseInsensitiveCollisions(t *testing.T) {
	manifest := fixtureManifest("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
	manifest.Projects[1] = manifest.Projects[0]
	manifest.Projects[1].ID = "two"
	manifest.Projects[1].Path = "projects/third_party/ONE"
	manifest.Projects[1].Source.Path = "examples/third_party/ONE"
	manifest.Projects[1].Notes = "changed"
	if err := ValidateManifest(manifest); err == nil {
		t.Fatal("ValidateManifest accepted case-insensitive project collision")
	}
}

func TestSyncLocalCheckoutIsDeterministicAndRemovesStaleFiles(t *testing.T) {
	upstream, commit := createGitFixture(t)
	corpusRoot := t.TempDir()
	manifest := fixtureManifest(commit)
	manifestPath := filepath.Join(corpusRoot, "manifest.json")
	writeManifest(t, manifestPath, manifest)

	if err := Sync(context.Background(), SyncOptions{ManifestPath: manifestPath, CorpusRoot: corpusRoot, UpstreamCheckout: upstream}); err != nil {
		t.Fatal(err)
	}
	target := filepath.Join(corpusRoot, "projects", "third_party")
	if _, err := os.Stat(filepath.Join(target, "one", "LICENSE")); err != nil {
		t.Fatal(err)
	}
	mainContents, err := os.ReadFile(filepath.Join(target, "one", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	if string(mainContents) != "Attribute VB_Name = \"Main\"\r\n" {
		t.Fatalf("copied project content changed: %q", mainContents)
	}
	afterFirstSync, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(target, "stale.bas"), []byte("stale"), 0o644); err != nil {
		t.Fatal(err)
	}
	withStale, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	if withStale == afterFirstSync {
		t.Fatal("digest unexpectedly ignored stale file")
	}
	if err := Sync(context.Background(), SyncOptions{ManifestPath: manifestPath, CorpusRoot: corpusRoot, UpstreamCheckout: upstream}); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(target, "stale.bas")); !os.IsNotExist(err) {
		t.Fatalf("stale file still exists, stat error = %v", err)
	}
	afterSecondSync, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	if afterFirstSync != afterSecondSync {
		t.Fatalf("resync changed the managed tree: %s != %s", afterFirstSync, afterSecondSync)
	}
	third, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	if afterSecondSync != third {
		t.Fatalf("second sync was not stable: %s != %s", afterSecondSync, third)
	}
}

func TestSyncRejectsCommitMismatchWithoutChangingExistingTree(t *testing.T) {
	upstream, commit := createGitFixture(t)
	corpusRoot := t.TempDir()
	manifest := fixtureManifest(commit)
	manifestPath := filepath.Join(corpusRoot, "manifest.json")
	writeManifest(t, manifestPath, manifest)
	if err := Sync(context.Background(), SyncOptions{ManifestPath: manifestPath, CorpusRoot: corpusRoot, UpstreamCheckout: upstream}); err != nil {
		t.Fatal(err)
	}
	target := filepath.Join(corpusRoot, "projects", "third_party")
	before, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	manifest.Upstream.Commit = strings.Repeat("0", 40)
	writeManifest(t, manifestPath, manifest)
	if err := Sync(context.Background(), SyncOptions{ManifestPath: manifestPath, CorpusRoot: corpusRoot, UpstreamCheckout: upstream}); err == nil {
		t.Fatal("Sync accepted a checkout at the wrong commit")
	}
	after, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	if before != after {
		t.Fatalf("failed sync changed existing tree: %s != %s", before, after)
	}
}

func TestSyncRejectsDirtyManagedTree(t *testing.T) {
	upstream, commit := createGitFixture(t)
	corpusRoot := t.TempDir()
	for _, args := range [][]string{{"init", "--quiet"}, {"config", "user.email", "test@example.invalid"}, {"config", "user.name", "xlflow test"}, {"commit", "--quiet", "--allow-empty", "-m", "baseline"}} {
		if err := runGit(t.Context(), corpusRoot, args...); err != nil {
			t.Fatal(err)
		}
	}
	manifestPath := filepath.Join(corpusRoot, "manifest.json")
	writeManifest(t, manifestPath, fixtureManifest(commit))
	target := filepath.Join(corpusRoot, "projects", "third_party")
	if err := os.MkdirAll(target, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(target, "local-change.bas"), []byte("dirty"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := Sync(context.Background(), SyncOptions{ManifestPath: manifestPath, CorpusRoot: corpusRoot, UpstreamCheckout: upstream}); err == nil || !strings.Contains(err.Error(), "uncommitted changes") {
		t.Fatalf("Sync() error = %v, want dirty-tree rejection", err)
	}
}

func TestPublishManagedTreeRestoresPreviousTreeAfterRenameFailure(t *testing.T) {
	root := t.TempDir()
	staged := filepath.Join(root, "stage", "third_party")
	target := filepath.Join(root, "projects", "third_party")
	if err := os.MkdirAll(staged, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(target, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(target, "old.bas"), []byte("old"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(staged, "new.bas"), []byte("new"), 0o644); err != nil {
		t.Fatal(err)
	}
	before, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	failPublish := func(old, new string) error {
		if strings.HasSuffix(filepath.ToSlash(old), "/stage/third_party") {
			return os.ErrPermission
		}
		return os.Rename(old, new)
	}
	if err := publishManagedTreeWith(staged, target, root, failPublish); err == nil {
		t.Fatal("publishManagedTreeWith unexpectedly succeeded")
	}
	after, err := TreeDigest(target)
	if err != nil {
		t.Fatal(err)
	}
	if before != after {
		t.Fatalf("rollback changed previous tree: %s != %s", before, after)
	}
	if _, err := os.Stat(filepath.Join(target, "old.bas")); err != nil {
		t.Fatalf("old tree was not restored: %v", err)
	}
}

func fixtureManifest(commit string) Manifest {
	return Manifest{
		SchemaVersion: SchemaVersion,
		Upstream:      Upstream{Repository: "harumiWeb/tree-sitter-vba", Commit: commit},
		Projects: []Project{
			{ID: "one", Path: "projects/third_party/one", Profile: ProfileGenericVBA, Enabled: true, Source: Source{Origin: OriginTreeSitterVBA, Path: "examples/third_party/one"}, Provenance: Provenance{Repository: "https://example.test/one", License: "MIT", LicenseFile: "LICENSE", SourceFile: "SOURCE.md"}},
			{ID: "two", Path: "projects/third_party/two", Profile: ProfileExcel, Enabled: false, Notes: "fixture disabled in this test", Source: Source{Origin: OriginTreeSitterVBA, Path: "examples/third_party/two"}, Provenance: Provenance{Repository: "https://example.test/two", License: "MIT", LicenseFile: "LICENSE", SourceFile: "SOURCE.md"}},
		},
	}
}

func createGitFixture(t *testing.T) (string, string) {
	t.Helper()
	root := t.TempDir()
	for _, project := range []string{"one", "two"} {
		dir := filepath.Join(root, "examples", "third_party", project)
		if err := os.MkdirAll(dir, 0o755); err != nil {
			t.Fatal(err)
		}
		for name, body := range map[string]string{"Main.bas": "Attribute VB_Name = \"Main\"\r\n", "LICENSE": "MIT\n", "SOURCE.md": "source\n"} {
			if err := os.WriteFile(filepath.Join(dir, name), []byte(body), 0o644); err != nil {
				t.Fatal(err)
			}
		}
	}
	for _, args := range [][]string{{"init", "--quiet"}, {"config", "user.email", "test@example.invalid"}, {"config", "user.name", "xlflow test"}, {"add", "."}, {"commit", "--quiet", "-m", "fixture"}} {
		if err := runGit(t.Context(), root, args...); err != nil {
			t.Fatal(err)
		}
	}
	commit, err := runGitOutput(t.Context(), root, "rev-parse", "HEAD")
	if err != nil {
		t.Fatal(err)
	}
	return root, strings.TrimSpace(commit)
}

func writeManifest(t *testing.T, path string, manifest Manifest) {
	t.Helper()
	body, err := json.MarshalIndent(manifest, "", "  ")
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, append(body, '\n'), 0o644); err != nil {
		t.Fatal(err)
	}
}
