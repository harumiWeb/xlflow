package lspserver

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestCheckLoadsBuiltinDatabase(t *testing.T) {
	if err := Check(Options{RootDir: t.TempDir(), Config: config.Default()}); err != nil {
		t.Fatal(err)
	}
}

func TestFileURIPathRoundTripWithEscapedJapanesePath(t *testing.T) {
	path := filepath.Join(t.TempDir(), "日本 語#%dir", "Main.bas")
	uri := pathToFileURI(path)
	if !strings.HasPrefix(uri, "file:") {
		t.Fatalf("uri = %q, want file URI", uri)
	}

	got, err := fileURIToPath(uri)
	if err != nil {
		t.Fatal(err)
	}
	if normalizePathKey(got) != normalizePathKey(path) {
		t.Fatalf("roundtrip path = %q, want %q via %q", got, path, uri)
	}
}

func TestDocumentsOverlayUsesUnsavedChangesAndClearsOnClose(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	docs := newDocuments(root)

	opened, err := docs.open(uri, "Option Explicit\nSub OldName()\nEnd Sub\n")
	if err != nil {
		t.Fatal(err)
	}
	if opened.ModuleKind != "standard" {
		t.Fatalf("module kind = %q, want standard", opened.ModuleKind)
	}

	changed, err := docs.change(uri, "Option Explicit\nSub NewName()\nEnd Sub\n")
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(changed.Source, "NewName") {
		t.Fatalf("changed source was not stored: %q", changed.Source)
	}
	if len(docs.openDocuments()) != 1 {
		t.Fatalf("open documents = %d, want 1", len(docs.openDocuments()))
	}

	docs.close(uri)
	if len(docs.openDocuments()) != 0 {
		t.Fatalf("open documents = %d, want 0 after close", len(docs.openDocuments()))
	}
}

func TestNewCreatesLogFileWithoutStartingServer(t *testing.T) {
	root := t.TempDir()
	logPath := filepath.Join(".xlflow", "lsp.log")
	_, cleanup, err := New(Options{RootDir: root, Config: config.Default(), LogFile: logPath})
	if err != nil {
		t.Fatal(err)
	}
	cleanup()

	if _, err := os.Stat(filepath.Join(root, logPath)); err != nil {
		t.Fatal(err)
	}
}
