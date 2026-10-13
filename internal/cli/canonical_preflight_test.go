package cli

import (
	"bytes"
	"context"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestCanonicalFormPreflightExcludesPictureWithCodeSuffix(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	makePackSourceRoots(t, root, cfg)
	if err := config.Write(filepath.Join(root, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	raw := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/assets/Login.bas"}}]}`)
	writePackSourceModule(t, root, filepath.Join("src", "forms", "specs", "Login.json"), raw)
	writePackSourceModule(t, root, filepath.Join("src", "forms", "code", "Login.bas"), []byte("Option Explicit\n"))
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	writePackSourceModule(t, root, filepath.Join("src", "forms", "assets", "Login.bas"), bmp)
	spec, err := forms.ParseFormSpec(forms.SpecInput{Format: "json", DisplayPath: "src/forms/specs/Login.json"}, raw)
	if err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: root, stdout: new(bytes.Buffer), stderr: new(bytes.Buffer), json: true}
	if err := a.runFormWritePreflight(context.Background(), "form build", cfg, formWriteCommandOptions{Spec: spec}); err != nil {
		t.Fatalf("form build preflight treated picture as code: %v", err)
	}
	if stdout, err := runBuildCommandForTest(root, "--json", "lint"); err != nil {
		t.Fatalf("lint treated picture as code: %v\n%s", err, stdout)
	}
	// File pull replaces picture references with content-addressed paths but
	// deliberately retains the former resource inside the reserved directory.
	withoutPicture := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[]}`)
	if err := os.WriteFile(filepath.Join(root, "src", "forms", "specs", "Login.json"), withoutPicture, 0o644); err != nil {
		t.Fatal(err)
	}
	if stdout, err := runBuildCommandForTest(root, "--json", "lint"); err != nil {
		t.Fatalf("lint treated retained unreferenced asset as code: %v\n%s", err, stdout)
	}
}
