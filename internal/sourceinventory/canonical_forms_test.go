package sourceinventory

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

const canonicalTestSpec = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[]}`

func TestCanonicalFormDiscoveryRejectsAmbiguousInputs(t *testing.T) {
	for _, test := range []struct {
		name  string
		files map[string]string
		mode  string
	}{
		{"duplicate specs", map[string]string{"specs/Login.json": canonicalTestSpec, "specs/Login.yaml": canonicalTestSpec}, "sidecar"},
		{"wrong form name", map[string]string{"specs/Login.json": strings.Replace(canonicalTestSpec, "Login", "Other", 1)}, "sidecar"},
		{"orphan code", map[string]string{"code/Login.bas": "Option Explicit"}, "sidecar"},
		{"nested code", map[string]string{"specs/Login.json": canonicalTestSpec, "code/nested/Login.bas": "Option Explicit"}, "sidecar"},
		{"sidecar attributes", map[string]string{"specs/Login.json": canonicalTestSpec, "code/Login.bas": `Attribute VB_Name = "Login"`}, "sidecar"},
		{"frm required", map[string]string{"specs/Login.json": canonicalTestSpec}, "frm"},
		{"fake attribute in comment", map[string]string{"specs/Login.json": canonicalTestSpec, "Login.frm": `' Attribute VB_Name = "Login"`}, "frm"},
	} {
		t.Run(test.name, func(t *testing.T) {
			root := t.TempDir()
			cfg := config.Default()
			cfg.UserForm.CodeSource = test.mode
			for path, body := range test.files {
				writeInventoryFile(t, root, "src/forms", path, body)
			}
			if _, err := collectCanonicalForms(root, filepath.Join(root, "src/forms"), Options{Config: cfg}); err == nil {
				t.Fatal("accepted invalid canonical layout")
			}
		})
	}
}

func TestCanonicalFormsIgnoreUnusedDesignerAndAllowExternalRoots(t *testing.T) {
	root := t.TempDir()
	base := t.TempDir()
	cfg := config.Default()
	writeInventoryFile(t, base, "specs", "Login.json", canonicalTestSpec)
	writeInventoryFile(t, base, "code", "Login.bas", "Option Explicit")
	// The stale compatibility designer is deliberately not authoring input.
	writeInventoryFile(t, base, "", "Login.frm", "stale Designer bytes")
	writeInventoryFile(t, base, "", "Login.frx", "opaque resource")
	items, err := collectCanonicalForms(root, base, Options{Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || string(items[0].Source) != "Option Explicit" {
		t.Fatalf("items=%+v", items)
	}
	if _, err := os.Stat(filepath.Join(base, "Login.frm")); err != nil {
		t.Fatal(err)
	}
}

func TestCanonicalLegacyFormsRemainRecursiveAndAttachSidecars(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	frm := "VERSION 5.00\nAttribute VB_Name = \"Legacy\"\nOption Explicit\n"
	writeInventoryFile(t, root, "src/forms/legacy", "Legacy.frm", frm)
	writeInventoryFile(t, root, "src/forms/code", "Legacy.bas", "Option Explicit\nPublic Sub Run()\nEnd Sub\n")
	items, err := collectCanonicalForms(root, base, Options{Config: cfg, AllowLegacyFormArtifacts: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || items[0].FormSpec != nil || !strings.HasSuffix(items[0].SourcePath, "legacy/Legacy.frm") {
		t.Fatalf("items = %+v", items)
	}
	if _, ok := items[0].Artifact("src/forms/code/Legacy.bas"); !ok {
		t.Fatalf("legacy sidecar was not attached: %+v", items[0].RelatedPaths())
	}
}

func TestCanonicalLegacySidecarRejectsAmbiguousNestedForms(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	frm := "VERSION 5.00\nAttribute VB_Name = \"Legacy\"\n"
	writeInventoryFile(t, root, "src/forms/one", "Legacy.frm", frm)
	writeInventoryFile(t, root, "src/forms/two", "Legacy.frm", frm)
	writeInventoryFile(t, root, "src/forms/code", "Legacy.bas", "Option Explicit\n")
	if _, err := collectCanonicalForms(root, base, Options{Config: cfg, AllowLegacyFormArtifacts: true}); err == nil {
		t.Fatal("accepted sidecar for ambiguous nested legacy forms")
	}
}

func TestCanonicalFormsRejectEmptyNestedDirectory(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Join(base, "empty"), 0o755); err != nil {
		t.Fatal(err)
	}
	if _, err := collectCanonicalForms(root, base, Options{Config: config.Default()}); err == nil {
		t.Fatal("accepted empty nested canonical form directory")
	}
}

func TestCanonicalFormsReportFirstInvalidDirectoryDeterministically(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	for _, name := range []string{"Zulu", "Alpha"} {
		if err := os.MkdirAll(filepath.Join(base, name), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	for range 64 {
		_, err := collectCanonicalForms(root, base, Options{Config: config.Default()})
		if err == nil || !strings.Contains(err.Error(), "src/forms/Alpha") {
			t.Fatalf("expected first directory Alpha, got %v", err)
		}
	}
}

func TestCanonicalFormsReportFirstAmbiguousLegacyFormDeterministically(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	for _, name := range []string{"Zulu", "Alpha"} {
		for _, dir := range []string{"one", "two"} {
			writeInventoryFile(t, root, "src/forms/"+dir, name+".frm", "VERSION 5.00\n")
		}
		writeInventoryFile(t, root, "src/forms/code", name+".bas", "Option Explicit\n")
	}
	for range 64 {
		_, err := collectCanonicalForms(root, base, Options{Config: config.Default(), AllowLegacyFormArtifacts: true})
		if err == nil || !strings.Contains(err.Error(), "ambiguous legacy form Alpha:") {
			t.Fatalf("expected first ambiguous form Alpha, got %v", err)
		}
	}
}

func TestCanonicalFormSpecPreservesAuthoredRootOmission(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	body := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login","observed":{"caption":"Observed","width":300,"height":200}},"controls":[{"id":"label1","name":"Label1","type":"Label","caption":"Authored"}]}`
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", body)
	cfg := config.Default()
	items, err := collectCanonicalForms(root, base, Options{Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || items[0].FormSpec == nil {
		t.Fatalf("items = %+v", items)
	}
	spec := items[0].FormSpec
	if spec.Form.Caption != nil || spec.Form.Width != nil || spec.Form.Height != nil || spec.Form.Build != nil {
		t.Fatalf("inferred root fields survived: %+v", spec.Form)
	}
	if spec.Form.Observed == nil || spec.Form.Observed.Width == nil || *spec.Form.Observed.Width != 300 {
		t.Fatalf("explicit observed root fields were lost: %+v", spec.Form.Observed)
	}
	if len(spec.Controls) != 1 || spec.Controls[0].Caption == nil || *spec.Controls[0].Caption != "Authored" || spec.Controls[0].Observed != nil {
		t.Fatalf("control authored fields = %+v", spec.Controls)
	}
}

func TestCanonicalFormsRetainReservedAssetsAndResolveNumericPictureNames(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/123abc456def.bmp"}}]}`)
	writeInventoryFile(t, root, "src/forms/code", "Login.bas", "Private Sub UserForm_Click()\nEnd Sub\n")
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(base, "assets"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(base, "123abc456def.bmp"), bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	// Pull retains unreferenced content-addressed files. The reserved folder
	// accepts arbitrary retained user files and is not scanned as components.
	if err := os.WriteFile(filepath.Join(base, "assets", "9f0012ab.bmp"), bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(base, "assets", "metadata.txt"), []byte("retained"), 0o644); err != nil {
		t.Fatal(err)
	}

	items, err := collectCanonicalForms(root, base, Options{
		Config: cfg, IgnoreCanonicalCompatibilityArtifacts: true, PreserveMissingFormCode: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || items[0].FormSpec == nil || len(items[0].FormSpec.Controls) != 1 {
		t.Fatalf("items = %+v", items)
	}
	if got := items[0].FormSpec.Controls[0].Picture.Data; !bytes.Equal(got, bmp) {
		t.Fatal("numeric-named referenced BMP was not resolved as raw image bytes")
	}
	if len(items[0].Related) != 3 { // canonical spec, sidecar, referenced BMP
		t.Fatalf("related artifacts = %+v", items[0].RelatedPaths())
	}
	var pictureRoleFound bool
	for _, artifact := range items[0].Related {
		if artifact.Role == ArtifactRolePicture {
			pictureRoleFound = true
		}
		if strings.Contains(artifact.Path, "/assets/") {
			t.Fatalf("unreferenced retained asset became an input: %+v", artifact)
		}
	}
	if !pictureRoleFound {
		t.Fatal("referenced picture artifact has no explicit picture role")
	}
}

func TestCanonicalPictureRoleOverridesModuleSuffix(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/images/logo.bas"}}]}`)
	writeInventoryFile(t, root, "src/forms/code", "Login.bas", "Private Sub FormCode()\nEnd Sub\n")
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(base, "images"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(base, "images", "logo.bas"), bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	items, err := collectCanonicalForms(root, base, Options{Config: cfg, PreserveMissingFormCode: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || string(items[0].Source) != "Private Sub FormCode()\nEnd Sub\n" {
		t.Fatalf("picture bytes were selected as form code: %+v", items)
	}
	if len(items[0].Related) != 3 {
		t.Fatalf("related artifacts = %+v", items[0].RelatedPaths())
	}
	for _, artifact := range items[0].Related {
		if strings.HasSuffix(strings.ToLower(artifact.Path), "logo.bas") && artifact.Role != ArtifactRolePicture {
			t.Fatalf("module-suffixed picture has role %q", artifact.Role)
		}
	}
}

func TestCanonicalPictureRoleDoesNotSelectCodeSidecar(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/code/Login.bas"}}]}`)
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(base, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(base, "code", "Login.bas"), bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	items, err := collectCanonicalForms(root, base, Options{Config: cfg, PreserveMissingFormCode: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || !items[0].PreserveExistingCode || len(items[0].Source) != 0 {
		t.Fatalf("picture bytes were selected as code sidecar: %+v", items)
	}
	if len(items[0].Related) != 2 {
		t.Fatalf("related artifacts = %+v", items[0].RelatedPaths())
	}
	if items[0].Related[1].Role != ArtifactRolePicture || !strings.HasSuffix(items[0].Related[1].Path, "src/forms/code/Login.bas") {
		t.Fatalf("code-directory picture artifact = %+v", items[0].Related[1])
	}
}

func TestCanonicalFormsRejectUnrelatedExtensionOutsideReservedAssets(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(base, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(base, "notes.txt"), []byte("not a form artifact"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := collectCanonicalForms(root, base, Options{Config: config.Default()}); err == nil {
		t.Fatal("accepted unrelated extension under forms root")
	}
}

func TestCanonicalFormsIgnoreCompatibilityFilesBeforeReadingInSidecarFileMode(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", canonicalTestSpec)
	writeInventoryFile(t, root, "src/forms/code", "Login.bas", "Option Explicit\n")
	writeInventoryFile(t, root, "src/forms", "Login.frm", "stale and must not be read")
	writeInventoryFile(t, root, "src/forms", "Login.frx", "stale binary")
	items, err := collectCanonicalForms(root, base, Options{
		Config: cfg, IgnoreCanonicalCompatibilityArtifacts: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 {
		t.Fatalf("items = %+v", items)
	}
	for _, path := range items[0].RelatedPaths() {
		if strings.EqualFold(filepath.Ext(path), ".frm") || strings.EqualFold(filepath.Ext(path), ".frx") {
			t.Fatalf("stale compatibility artifact was captured: %v", items[0].RelatedPaths())
		}
	}
}

func TestCanonicalFormsRejectStaleFRMWhenItWouldProvideCode(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	staleSpec := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[],"warnings":[{"code":"compatibility_artifact_unsynchronized","message":"stale"}]}`
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", staleSpec)
	writeInventoryFile(t, root, "src/forms", "Login.frm", "Attribute VB_Name = \"Login\"\nOption Explicit\n")
	for _, mode := range []string{"frm", "sidecar"} {
		t.Run(mode, func(t *testing.T) {
			cfg := config.Default()
			cfg.UserForm.CodeSource = mode
			_, err := collectCanonicalForms(root, base, Options{Config: cfg, AllowLegacyFormArtifacts: true})
			if err == nil || !strings.Contains(err.Error(), "FRM201") {
				t.Fatalf("error = %v, want FRM201", err)
			}
		})
	}
}
