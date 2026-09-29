package sourceinventory

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDiscoverIsDeterministicAndCapturesFormArtifacts(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	writeInventoryFile(t, root, cfg.Src.Modules, "Zed.bas", "Attribute VB_Name = \"Zed\"\r\n")
	writeInventoryFile(t, root, cfg.Src.Classes, "Alpha.cls", "VERSION 1.0 CLASS\r\nEND\r\nAttribute VB_Name = \"Alpha\"\r\nAttribute VB_Exposed = False\r\n")
	writeInventoryFile(t, root, cfg.Src.Workbook, "ThisWorkbook.bas", "Option Explicit\r\n")
	writeInventoryFile(t, root, cfg.Src.Forms, "Login.frm", "Attribute VB_Name = \"Login\"\r\n")
	writeInventoryFile(t, root, cfg.Src.Forms, "Login.frx", "binary")
	writeInventoryFile(t, root, filepath.Join(cfg.Src.Forms, "code"), "Login.bas", "Option Explicit\r\n")

	before := inventoryTree(t, root)
	first, err := Discover(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	second, err := Discover(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("inventory is not deterministic:\nfirst=%#v\nsecond=%#v", first, second)
	}
	if after := inventoryTree(t, root); !reflect.DeepEqual(before, after) {
		t.Fatal("discovery modified source artifacts")
	}
	if got := []string{first[0].SourcePath, first[1].SourcePath, first[2].SourcePath, first[3].SourcePath}; !reflect.DeepEqual(got, []string{
		"src/classes/Alpha.cls", "src/forms/Login.frm", "src/modules/Zed.bas", "src/workbook/ThisWorkbook.bas",
	}) {
		t.Fatalf("paths = %#v", got)
	}
	if got := first[1].RelatedPaths(); !reflect.DeepEqual(got, []string{"src/forms/Login.frx", "src/forms/code/Login.bas"}) {
		t.Fatalf("form artifacts = %#v", got)
	}
}

func TestDiscoverRejectsMissingRootsAndUnsupportedFiles(t *testing.T) {
	t.Run("missing configured root", func(t *testing.T) {
		root := t.TempDir()
		cfg := config.Default()
		_, err := Discover(Options{Root: root, Config: cfg})
		if err == nil || !strings.Contains(err.Error(), "read standard source root") {
			t.Fatalf("err = %v", err)
		}
	})
	t.Run("unsupported extension", func(t *testing.T) {
		root := t.TempDir()
		cfg := config.Default()
		for _, dir := range []string{cfg.Src.Modules, cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook} {
			if err := os.MkdirAll(filepath.Join(root, filepath.FromSlash(dir)), 0o755); err != nil {
				t.Fatal(err)
			}
		}
		writeInventoryFile(t, root, cfg.Src.Modules, "notes.txt", "not VBA")
		_, err := Discover(Options{Root: root, Config: cfg})
		if err == nil || !strings.Contains(err.Error(), "unsupported standard source file") {
			t.Fatalf("err = %v", err)
		}
	})
}

func TestDiscoverAllowsAbsolutePackRoots(t *testing.T) {
	root := t.TempDir()
	external := t.TempDir()
	cfg := config.Default()
	cfg.Src.Modules = filepath.Join(external, "modules")
	cfg.Src.Classes = filepath.Join(external, "classes")
	cfg.Src.Forms = filepath.Join(external, "forms")
	cfg.Src.Workbook = filepath.Join(external, "workbook")
	for _, dir := range []string{cfg.Src.Modules, cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook} {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			t.Fatal(err)
		}
	}
	writeInventoryFile(t, "", cfg.Src.Modules, "External.bas", "Attribute VB_Name = \"External\"\r\n")
	components, err := Discover(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(components) != 1 || !filepath.IsAbs(components[0].SourcePath) {
		t.Fatalf("components = %#v", components)
	}
}

func writeInventoryFile(t *testing.T, root, dir, name, body string) {
	t.Helper()
	path := filepath.Join(root, filepath.FromSlash(dir), name)
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func inventoryTree(t *testing.T, root string) map[string]string {
	t.Helper()
	out := map[string]string{}
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil || d.IsDir() {
			return err
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		rel, _ := filepath.Rel(root, path)
		out[filepath.ToSlash(rel)] = string(body)
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	return out
}
