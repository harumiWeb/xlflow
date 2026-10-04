package filepull

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/picture"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

func TestPullExtractsSupportedModulesAndReconcilesManagedFiles(t *testing.T) {
	root := t.TempDir()
	cfg := testConfig()
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	stale := filepath.Join(root, "src", "modules", "Stale.bas")
	writeTestFile(t, stale, "stale")
	unmanaged := filepath.Join(root, "src", "modules", "README.md")
	writeTestFile(t, unmanaged, "keep")
	form := filepath.Join(root, "src", "forms", "Old.frm")
	writeTestFile(t, form, "keep form")

	result, err := Pull(root, cfg, workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Modules.Standard != 1 || result.Modules.Class != 1 || result.Modules.Document != 2 || result.CodePage != 932 {
		t.Fatalf("unexpected result: %+v", result)
	}
	for _, path := range []string{
		filepath.Join(root, "src", "modules", "Module1.bas"),
		filepath.Join(root, "src", "classes", "Class1.cls"),
		filepath.Join(root, "src", "workbook", "Sheet1.bas"),
		filepath.Join(root, "src", "workbook", "ThisWorkbook.bas"),
	} {
		body, readErr := os.ReadFile(path)
		if readErr != nil {
			t.Fatal(readErr)
		}
		if bytes.HasPrefix(body, []byte{0xef, 0xbb, 0xbf}) || !bytes.HasSuffix(body, []byte("\n")) {
			t.Fatalf("%s is not UTF-8 without BOM with final newline", path)
		}
	}
	classBody, err := os.ReadFile(filepath.Join(root, "src", "classes", "Class1.cls"))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.HasPrefix(string(classBody), "VERSION 1.0 CLASS\nBEGIN\n") || strings.Contains(string(classBody), "Attribute VB_Base") {
		t.Fatalf("unexpected class export:\n%s", classBody)
	}
	documentBody, err := os.ReadFile(filepath.Join(root, "src", "workbook", "Sheet1.bas"))
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(documentBody), "Attribute VB_") || strings.HasPrefix(string(documentBody), "VERSION") {
		t.Fatalf("unexpected document export:\n%s", documentBody)
	}
	if _, err := os.Stat(stale); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("stale managed source remains: %v", err)
	}
	for _, path := range []string{unmanaged, form} {
		if _, err := os.Stat(path); err != nil {
			t.Fatalf("unmanaged source changed: %s: %v", path, err)
		}
	}
}

func TestPullExtractsUserFormSidecarsAndSpecs(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readFixture(t, "p4_form.bin"))
	result, err := Pull(root, testConfig(), workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Modules.Form != 1 {
		t.Fatalf("form count = %d, want 1", result.Modules.Form)
	}
	codePath := filepath.Join(root, "src", "forms", "code", "UserForm1.bas")
	code, err := os.ReadFile(codePath)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(code), "Attribute VB_") || !bytes.HasSuffix(code, []byte("\n")) {
		t.Fatalf("unexpected code sidecar:\n%s", code)
	}
	specPath := filepath.Join(root, "src", "forms", "specs", "UserForm1.yaml")
	specBody, err := os.ReadFile(specPath)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"kind: xlflow.userform", "name: UserForm1", "controls:", forms.CompatibilityArtifactUnsynchronizedWarningCode} {
		if !bytes.Contains(specBody, []byte(want)) {
			t.Fatalf("spec missing %q:\n%s", want, specBody)
		}
	}
}

func TestPullExtractsNestedUserFormAndReconcilesManagedArtifacts(t *testing.T) {
	root := t.TempDir()
	formsRoot := filepath.Join(root, "src", "forms")
	for path, body := range map[string]string{
		filepath.Join(formsRoot, "specs", "OldForm.json"):   "stale",
		filepath.Join(formsRoot, "specs", "UserForm1.json"): "stale format",
		filepath.Join(formsRoot, "code", "OldForm.bas"):     "stale",
		filepath.Join(formsRoot, "Keep.frm"):                "generated compatibility artifact",
		filepath.Join(formsRoot, "specs", "README.md"):      "unmanaged",
	} {
		writeTestFile(t, path, body)
	}
	workbook := writeWorkbook(t, root, readFixture(t, "p6_nested_form.bin"))
	result, err := Pull(root, testConfig(), workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Modules.Form != 1 {
		t.Fatalf("form count = %d, want 1", result.Modules.Form)
	}
	specPath := filepath.Join(formsRoot, "specs", "UserForm1.yaml")
	specBody, err := os.ReadFile(specPath)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"type: Frame", "type: MultiPage", "parentId:"} {
		if !bytes.Contains(specBody, []byte(want)) {
			t.Fatalf("nested spec missing %q:\n%s", want, specBody)
		}
	}
	for _, stale := range []string{
		filepath.Join(formsRoot, "specs", "OldForm.json"),
		filepath.Join(formsRoot, "specs", "UserForm1.json"),
		filepath.Join(formsRoot, "code", "OldForm.bas"),
	} {
		if _, err := os.Stat(stale); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("stale managed form artifact remains: %s: %v", stale, err)
		}
	}
	for _, kept := range []string{filepath.Join(formsRoot, "Keep.frm"), filepath.Join(formsRoot, "specs", "README.md")} {
		if _, err := os.Stat(kept); err != nil {
			t.Fatalf("unmanaged form artifact changed: %s: %v", kept, err)
		}
	}
}

func TestPullExtractsExcelAuthoredImageAssetsRecursively(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readPictureFixture(t, "baseline.bin"))
	result, err := Pull(root, testConfig(), workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Modules.Form != 1 {
		t.Fatalf("form count = %d, want 1", result.Modules.Form)
	}

	specs, err := os.ReadDir(filepath.Join(root, "src", "forms", "specs"))
	if err != nil {
		t.Fatal(err)
	}
	if len(specs) != 1 {
		t.Fatalf("spec count = %d, want 1", len(specs))
	}
	specBody, err := os.ReadFile(filepath.Join(root, "src", "forms", "specs", specs[0].Name()))
	if err != nil {
		t.Fatal(err)
	}
	var spec forms.FormSpec
	if err := yaml.Unmarshal(specBody, &spec); err != nil {
		t.Fatalf("decode pulled FormSpec: %v\n%s", err, specBody)
	}
	controls := make(map[string]forms.FormSpecControl, len(spec.Controls))
	for _, control := range spec.Controls {
		controls[control.Name] = control
	}
	frame, ok := controls["FrameMain"]
	if !ok || frame.Type != "Frame" {
		t.Fatalf("nested frame missing from pulled spec: %+v", frame)
	}
	assets := excelAuthoredPictureAssets(t)
	paths := make(map[string]string, 2)
	for _, name := range []string{"ImageBmp", "ImageJpeg"} {
		control, ok := controls[name]
		if !ok || control.Type != "Image" || control.Picture == nil {
			t.Fatalf("picture control %q missing from pulled spec: %+v", name, control)
		}
		if name == "ImageJpeg" && control.ParentID != frame.ID {
			t.Fatalf("nested image parent = %q, want FrameMain id %q", control.ParentID, frame.ID)
		}
		asset := assets[name]
		sum := sha256.Sum256(asset.Data)
		wantRelative := filepath.ToSlash(filepath.Join("src", "forms", "assets", hex.EncodeToString(sum[:])+".bmp"))
		if control.Picture.Path != wantRelative {
			t.Fatalf("%s picture path = %q, want %q", name, control.Picture.Path, wantRelative)
		}
		assetPath := filepath.Join(root, filepath.FromSlash(control.Picture.Path))
		paths[name] = control.Picture.Path
		assetBody, err := os.ReadFile(assetPath)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(assetBody, asset.Data) {
			t.Fatalf("%s asset differs from decoded Excel-authored picture bytes", name)
		}
		if !containsPath(result.Written, assetPath) {
			t.Fatalf("pull result does not report picture asset %s: %+v", assetPath, result.Written)
		}
		if slices.ContainsFunc(control.Unsupported, func(property string) bool { return strings.EqualFold(property, "picture") }) {
			t.Fatalf("successfully extracted picture remains unsupported on %s: %+v", name, control.Unsupported)
		}
		for _, warning := range spec.Warnings {
			if warning.Code != "unsupported_properties" || !strings.EqualFold(warning.Control, name) {
				continue
			}
			const warningPrefix = "Unsupported Designer properties were preserved only in the binary model: "
			properties, ok := strings.CutPrefix(warning.Message, warningPrefix)
			if !ok {
				continue
			}
			for _, property := range strings.Split(strings.TrimSuffix(properties, "."), ", ") {
				if strings.EqualFold(strings.TrimSpace(property), "picture") {
					t.Fatalf("successfully extracted picture remains in unsupported warning for %s: %+v", name, warning)
				}
			}
		}
	}
	if assets["ImageBmp"].Format != "bmp" || assets["ImageJpeg"].Format != "bmp" || assets["ImageBmp"].Width != 24 || assets["ImageBmp"].Height != 16 || assets["ImageJpeg"].Width != 24 || assets["ImageJpeg"].Height != 16 {
		t.Fatalf("Excel-authored fixture picture decode does not match environment evidence: %+v", assets)
	}
	if bytes.Equal(assets["ImageBmp"].Data, assets["ImageJpeg"].Data) || paths["ImageBmp"] == paths["ImageJpeg"] {
		t.Fatalf("Excel-authored BMP and JPEG inputs did not persist as distinct BMP assets: paths=%+v", paths)
	}
}

func TestExtractFormPicturesRequiresExactProjectedControlMapping(t *testing.T) {
	project, err := vbaproject.Read(readPictureFixture(t, "baseline.bin"))
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Forms) != 1 {
		t.Fatalf("Excel-authored fixture form count = %d, want 1", len(project.Forms))
	}
	source := project.Forms[0]
	for _, test := range []struct {
		name   string
		mutate func(*forms.FormSpecControl)
	}{
		{name: "name", mutate: func(control *forms.FormSpecControl) { control.Name += "Changed" }},
		{name: "type", mutate: func(control *forms.FormSpecControl) { control.Type = "Frame" }},
	} {
		t.Run(test.name, func(t *testing.T) {
			projected, err := projection.Project(source)
			if err != nil {
				t.Fatal(err)
			}
			projected.Controls = slices.Clone(projected.Controls)
			test.mutate(&projected.Controls[0])

			if _, _, err := extractFormPictures(source, &projected); err == nil || !strings.Contains(err.Error(), "does not match projected control") {
				t.Fatalf("control mapping error = %v, want name/type mismatch", err)
			}
		})
	}
}

func TestPullDeduplicatesRepeatedImageAssets(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readPictureFixture(t, "baseline.bin"))
	project, err := vbaproject.Read(readPictureFixture(t, "baseline.bin"))
	if err != nil {
		t.Fatal(err)
	}
	var imageBMP, imageJPEG *oforms.Control
	var visit func([]*oforms.Control)
	visit = func(controls []*oforms.Control) {
		for _, control := range controls {
			if control == nil {
				continue
			}
			switch control.Name {
			case "ImageBmp":
				imageBMP = control
			case "ImageJpeg":
				imageJPEG = control
			}
			children := control.Children
			if control.MultiPage != nil {
				children = control.MultiPage.Pages
			}
			visit(children)
		}
	}
	visit(project.Forms[0].Controls)
	if imageBMP == nil || imageJPEG == nil || imageBMP.Record == nil || imageJPEG.Record == nil {
		t.Fatalf("fixture Image controls were not found: BMP=%+v JPEG=%+v", imageBMP, imageJPEG)
	}
	imageJPEG.Record.Pictures["Picture"] = bytes.Clone(imageBMP.Record.Pictures["Picture"])
	p, err := buildProjectPlan(root, testConfig(), workbook, &inspectedProject{project: project})
	if err != nil {
		t.Fatal(err)
	}
	assetCount := 0
	var formSpec forms.FormSpec
	for _, file := range p.files {
		if file.preserveExistingAsset {
			assetCount++
		}
		if filepath.Ext(file.path) == ".yaml" && strings.Contains(filepath.ToSlash(file.path), "/forms/specs/") {
			if err := yaml.Unmarshal(file.body, &formSpec); err != nil {
				t.Fatal(err)
			}
		}
	}
	if assetCount != 1 {
		t.Fatalf("planned asset count = %d, want one deduplicated asset", assetCount)
	}
	paths := make(map[string]string)
	for _, control := range formSpec.Controls {
		if control.Name == "ImageBmp" || control.Name == "ImageJpeg" {
			if control.Picture == nil {
				t.Fatalf("picture missing from %s", control.Name)
			}
			paths[control.Name] = control.Picture.Path
		}
	}
	if paths["ImageBmp"] == "" || paths["ImageBmp"] != paths["ImageJpeg"] {
		t.Fatalf("duplicate picture paths differ: %+v", paths)
	}
}

func TestPullReusesExistingPictureAssetsAndKeepsUnmanagedAssets(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readPictureFixture(t, "baseline.bin"))
	first, err := Pull(root, testConfig(), workbook)
	if err != nil {
		t.Fatal(err)
	}
	assetsRoot := filepath.Join(root, "src", "forms", "assets")
	var assetPaths []string
	assetBodies := make(map[string][]byte)
	for _, path := range first.Written {
		if filepath.Dir(path) != assetsRoot {
			continue
		}
		body, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		assetPaths = append(assetPaths, path)
		assetBodies[path] = body
	}
	if len(assetPaths) != 2 {
		t.Fatalf("first pull asset count = %d, want 2: %+v", len(assetPaths), first.Written)
	}
	userAsset := filepath.Join(assetsRoot, "user-owned.txt")
	staleAsset := filepath.Join(assetsRoot, "previous-pull.bmp")
	writeTestFile(t, userAsset, "keep user asset")
	writeTestFile(t, staleAsset, "keep stale asset")

	originalPublish := publishArtifact
	t.Cleanup(func() { publishArtifact = originalPublish })
	assetWrites := 0
	publishArtifact = func(path string, body []byte, validate func(string) error) (coordination.PublishResult, error) {
		if _, exists := assetBodies[path]; exists {
			assetWrites++
		}
		return originalPublish(path, body, validate)
	}
	if _, err := Pull(root, testConfig(), workbook); err != nil {
		t.Fatal(err)
	}
	if assetWrites != 0 {
		t.Fatalf("identical picture assets were published %d times", assetWrites)
	}
	for path, want := range assetBodies {
		got, err := os.ReadFile(path)
		if err != nil || !bytes.Equal(got, want) {
			t.Fatalf("reused asset changed at %s: %v", path, err)
		}
	}
	for path, want := range map[string]string{userAsset: "keep user asset", staleAsset: "keep stale asset"} {
		got, err := os.ReadFile(path)
		if err != nil || string(got) != want {
			t.Fatalf("unmanaged asset changed at %s: %q, %v", path, got, err)
		}
	}
}

func TestPullPictureAssetCollisionFailsBeforePublication(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readPictureFixture(t, "baseline.bin"))
	initial, err := buildPlan(root, testConfig(), workbook)
	if err != nil {
		t.Fatal(err)
	}
	var conflictingAsset, specPath string
	for _, file := range initial.files {
		if file.preserveExistingAsset {
			conflictingAsset = file.path
		}
		if filepath.Ext(file.path) == ".yaml" && strings.Contains(filepath.ToSlash(file.path), "/forms/specs/") {
			specPath = file.path
		}
	}
	if conflictingAsset == "" || specPath == "" {
		t.Fatalf("fixture plan lacks asset/spec: asset=%q spec=%q", conflictingAsset, specPath)
	}
	if err := os.MkdirAll(filepath.Dir(conflictingAsset), 0o755); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, conflictingAsset, "user-owned conflict")
	writeTestFile(t, specPath, "keep existing spec")
	sentinel := filepath.Join(root, "src", "modules", "Existing.bas")
	writeTestFile(t, sentinel, "keep existing module")

	if _, err := Pull(root, testConfig(), workbook); !errors.Is(err, ErrPictureAssetConflict) {
		t.Fatalf("pull error = %v, want picture asset conflict", err)
	}
	for path, want := range map[string]string{
		conflictingAsset: "user-owned conflict",
		specPath:         "keep existing spec",
		sentinel:         "keep existing module",
	} {
		body, err := os.ReadFile(path)
		if err != nil || string(body) != want {
			t.Fatalf("file changed after picture collision at %s: %q, %v", path, body, err)
		}
	}
}

func TestPullPictureAssetPublicationFailureRollsBackAssets(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readPictureFixture(t, "baseline.bin"))
	initial, err := buildPlan(root, testConfig(), workbook)
	if err != nil {
		t.Fatal(err)
	}
	var assetPaths []string
	var specPath string
	for _, file := range initial.files {
		if file.preserveExistingAsset {
			assetPaths = append(assetPaths, file.path)
		}
		if filepath.Ext(file.path) == ".yaml" && strings.Contains(filepath.ToSlash(file.path), "/forms/specs/") {
			specPath = file.path
		}
	}
	if len(assetPaths) != 2 || specPath == "" {
		t.Fatalf("fixture plan lacks expected assets/spec: assets=%v spec=%q", assetPaths, specPath)
	}
	originalPublish := publishArtifact
	t.Cleanup(func() { publishArtifact = originalPublish })
	assetPublished := false
	publishArtifact = func(path string, body []byte, validate func(string) error) (coordination.PublishResult, error) {
		if path == specPath {
			return coordination.PublishResult{}, errors.New("injected FormSpec publication failure")
		}
		if strings.EqualFold(filepath.Dir(path), filepath.Dir(assetPaths[0])) {
			assetPublished = true
		}
		return originalPublish(path, body, validate)
	}
	if _, err := Pull(root, testConfig(), workbook); !errors.Is(err, ErrPublish) {
		t.Fatalf("pull error = %v, want publication failure", err)
	}
	if !assetPublished {
		t.Fatal("failure injection occurred before any picture asset was published")
	}
	for _, path := range assetPaths {
		if _, err := os.Lstat(path); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("picture asset remained after rollback at %s: %v", path, err)
		}
	}
}

func TestPullRejectsPictureAssetsEscapingProjectThroughFormsJunction(t *testing.T) {
	root := t.TempDir()
	outside := t.TempDir()
	formsPath := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Dir(formsPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(outside, formsPath); err != nil {
		t.Skipf("directory symlink/junction unavailable: %v", err)
	}
	workbook := writeWorkbook(t, root, readPictureFixture(t, "baseline.bin"))
	if _, err := Pull(root, testConfig(), workbook); !errors.Is(err, ErrUnsafeSourcePath) {
		t.Fatalf("pull error = %v, want unsafe source path", err)
	}
	if _, err := os.Stat(filepath.Join(outside, "assets")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("pull wrote picture assets outside project root: %v", err)
	}
}

func TestPullRejectsFrmUserFormBeforeMutation(t *testing.T) {
	root := t.TempDir()
	existing := filepath.Join(root, "src", "modules", "Existing.bas")
	writeTestFile(t, existing, "keep")
	workbook := writeWorkbook(t, root, readFixture(t, "p4_form.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "frm"
	_, err := Pull(root, cfg, workbook)
	if !errors.Is(err, ErrUserFormCodeSourceUnsupported) {
		t.Fatalf("error = %v", err)
	}
	body, readErr := os.ReadFile(existing)
	if readErr != nil || string(body) != "keep" {
		t.Fatalf("source mutated after rejected frm mode: %q, %v", body, readErr)
	}
}

func TestPullFrmWithoutUserFormsDoesNotManageSidecars(t *testing.T) {
	root := t.TempDir()
	existingSpec := filepath.Join(root, "src", "forms", "specs", "Existing.yaml")
	existingCode := filepath.Join(root, "src", "forms", "code", "Existing.bas")
	writeTestFile(t, existingSpec, "keep spec")
	writeTestFile(t, existingCode, "keep code")
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "frm"

	if _, err := Pull(root, cfg, workbook); err != nil {
		t.Fatal(err)
	}
	for path, want := range map[string]string{existingSpec: "keep spec", existingCode: "keep code"} {
		body, err := os.ReadFile(path)
		if err != nil || string(body) != want {
			t.Fatalf("frm sidecar changed: %s = %q, %v", path, body, err)
		}
	}
}

func TestPullRejectsMalformedUserFormBeforeMutation(t *testing.T) {
	root := t.TempDir()
	existing := filepath.Join(root, "src", "modules", "Existing.bas")
	writeTestFile(t, existing, "keep")
	malformed := replaceCFBStream(t, readFixture(t, "p4_form.bin"), "UserForm1/o", func(body []byte) []byte {
		return append(body, 0xff)
	})
	workbook := writeWorkbook(t, root, malformed)
	_, err := Pull(root, testConfig(), workbook)
	if !errors.Is(err, ErrUserFormDesignerMalformed) {
		t.Fatalf("error = %v", err)
	}
	body, readErr := os.ReadFile(existing)
	if readErr != nil || string(body) != "keep" {
		t.Fatalf("source mutated after malformed Designer: %q, %v", body, readErr)
	}
}

func TestExtractFormsRejectsUserFormIdentityMismatch(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readFixture(t, "p4_form.bin"))
	inspection, err := inspectProject(workbook)
	if err != nil {
		t.Fatal(err)
	}
	for i := range inspection.project.Modules {
		if inspection.project.Modules[i].Type == vbaproject.ModuleForm {
			inspection.project.Modules[i].Name = "RenamedForm"
		}
	}
	_, err = extractForms(inspection, testConfig())
	if !errors.Is(err, ErrUserFormIdentityMismatch) {
		t.Fatalf("error = %v", err)
	}
}

func TestProbeInspectsOnlyWorkbookCapabilityWithoutPublishing(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	result, err := Probe(workbook, testConfig())
	if err != nil {
		t.Fatal(err)
	}
	if !result.Supported || result.Reason != "supported" || result.CodePage != 932 || result.HasForms {
		t.Fatalf("probe result = %+v", result)
	}
	if _, err := os.Stat(filepath.Join(root, "src", "modules", "Module1.bas")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("probe published source: %v", err)
	}

	formWorkbook := writeWorkbook(t, root, readFixture(t, "p4_form.bin"))
	formResult, err := Probe(formWorkbook, testConfig())
	if err != nil {
		t.Fatal(err)
	}
	if !formResult.Supported || formResult.Reason != "supported" || !formResult.HasForms {
		t.Fatalf("form probe result = %+v", formResult)
	}
	frmConfig := testConfig()
	frmConfig.UserForm.CodeSource = "frm"
	frmResult, err := Probe(formWorkbook, frmConfig)
	if err != nil {
		t.Fatal(err)
	}
	if frmResult.Supported || frmResult.Reason != "userform_code_source" || !frmResult.HasForms {
		t.Fatalf("frm probe result = %+v", frmResult)
	}

	malformed := replaceCFBStream(t, readFixture(t, "p4_form.bin"), "UserForm1/o", func(body []byte) []byte {
		return append(body, 0xff)
	})
	malformedWorkbook := writeWorkbook(t, root, malformed)
	frmResult, err = Probe(malformedWorkbook, frmConfig)
	if err != nil {
		t.Fatalf("frm probe must route before Designer parsing: %v", err)
	}
	if frmResult.Supported || frmResult.Reason != "userform_code_source" || !frmResult.HasForms {
		t.Fatalf("malformed frm probe result = %+v", frmResult)
	}
	if _, err := Probe(malformedWorkbook, testConfig()); !errors.Is(err, ErrUserFormDesignerMalformed) {
		t.Fatalf("sidecar malformed probe error = %v, want ErrUserFormDesignerMalformed", err)
	}
}

func TestProbeLeavesSourceTreeValidationToPullContext(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	result, err := Probe(workbook, testConfig())
	if err != nil || !result.Supported {
		t.Fatalf("workbook capability probe = %+v, %v", result, err)
	}

	cfg := testConfig()
	cfg.Src.Forms = cfg.Src.Modules
	if _, err := Pull(root, cfg, workbook); !errors.Is(err, ErrUnsafeSourcePath) {
		t.Fatalf("Pull error = %v, want ErrUnsafeSourcePath", err)
	}
}

func BenchmarkProbeClosedWorkbook(b *testing.B) {
	root := b.TempDir()
	workbook := writeWorkbook(b, root, readFixture(b, "p1_compiled.bin"))
	b.ResetTimer()
	for b.Loop() {
		if _, err := Probe(workbook, testConfig()); err != nil {
			b.Fatal(err)
		}
	}
}

func TestPullRejectsProtectedAndMissingProjects(t *testing.T) {
	root := t.TempDir()
	protected := writeWorkbook(t, root, readFixture(t, "p3_protected.bin"))
	if _, err := Pull(root, testConfig(), protected); !errors.Is(err, ErrProtectedProject) {
		t.Fatalf("protected error = %v", err)
	}
	empty := filepath.Join(root, "empty.xlsm")
	writeZip(t, empty, nil)
	if _, err := Pull(root, testConfig(), empty); !errors.Is(err, ErrMissingVBAProject) {
		t.Fatalf("missing error = %v", err)
	}
}

func TestPullUsesFolderAnnotation(t *testing.T) {
	root := t.TempDir()
	project, err := vbaproject.Read(readFixture(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	for i := range project.Modules {
		if project.Modules[i].Type == vbaproject.ModuleStd {
			project.Modules[i].Source += "'@Folder(\"Domain.Services\")\r\n"
		}
	}
	bin, err := vbaproject.Write(project)
	if err != nil {
		t.Fatal(err)
	}
	cfg := testConfig()
	cfg.VBA.Folders = true
	cfg.VBA.FolderAnnotation = "update"
	workbook := writeWorkbook(t, root, bin)
	if _, err := Pull(root, cfg, workbook); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(root, "src", "modules", "Domain", "Services", "Module1.bas")); err != nil {
		t.Fatal(err)
	}
}

func TestPullReconcilesCaseOnlyModuleRename(t *testing.T) {
	root := t.TempDir()
	writeTestFile(t, filepath.Join(root, "src", "modules", "module1.bas"), "stale")
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	if _, err := Pull(root, testConfig(), workbook); err != nil {
		t.Fatal(err)
	}
	var matches []string
	err := filepath.WalkDir(filepath.Join(root, "src", "modules"), func(path string, entry os.DirEntry, err error) error {
		if err == nil && !entry.IsDir() && strings.EqualFold(entry.Name(), "Module1.bas") {
			matches = append(matches, path)
		}
		return err
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(matches) != 1 {
		t.Fatalf("case-only reconciliation left %d module files: %v", len(matches), matches)
	}
	body, err := os.ReadFile(matches[0])
	if err != nil || !strings.Contains(string(body), `Attribute VB_Name = "Module1"`) {
		t.Fatalf("case-only target was not replaced: %q, %v", body, err)
	}
}

func TestPullRejectsFormsRootOverlapBeforeMutation(t *testing.T) {
	tests := []struct {
		name      string
		configure func(*config.Config)
		formPath  string
	}{
		{"forms-inside-managed", func(cfg *config.Config) { cfg.Src.Forms = "src/modules/forms" }, filepath.Join("src", "modules", "forms", "keep.bas")},
		{"managed-inside-forms", func(cfg *config.Config) { cfg.Src.Forms = "src" }, filepath.Join("src", "forms", "keep.bas")},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			root := t.TempDir()
			cfg := testConfig()
			test.configure(&cfg)
			form := filepath.Join(root, test.formPath)
			writeTestFile(t, form, "keep")
			workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
			if _, err := Pull(root, cfg, workbook); !errors.Is(err, ErrUnsafeSourcePath) {
				t.Fatalf("error = %v, want unsafe source path", err)
			}
			body, err := os.ReadFile(form)
			if err != nil || string(body) != "keep" {
				t.Fatalf("forms root mutated after rejection: %q, %v", body, err)
			}
		})
	}
}

func TestPublishRollsBackEarlierFiles(t *testing.T) {
	root := t.TempDir()
	first := filepath.Join(root, "First.bas")
	second := filepath.Join(root, "Second.bas")
	writeTestFile(t, first, "old first")
	writeTestFile(t, second, "old second")
	p := plan{files: []plannedFile{{path: first, body: []byte("new first")}, {path: second, body: []byte("new second")}}}
	originalPublish := publishArtifact
	t.Cleanup(func() { publishArtifact = originalPublish })
	calls := 0
	publishArtifact = func(path string, body []byte, validate func(string) error) (coordination.PublishResult, error) {
		calls++
		if calls == 2 {
			return coordination.PublishResult{}, errors.New("injected publication failure")
		}
		return originalPublish(path, body, validate)
	}
	if err := publish(p); !errors.Is(err, ErrPublish) {
		t.Fatalf("error = %v", err)
	}
	for path, want := range map[string]string{first: "old first", second: "old second"} {
		body, err := os.ReadFile(path)
		if err != nil || string(body) != want {
			t.Fatalf("rollback %s = %q, %v", path, body, err)
		}
	}
}

func TestPublishRollbackRemovesNewDirectories(t *testing.T) {
	root := t.TempDir()
	target := filepath.Join(root, "src", "modules", "Nested", "Main.bas")
	originalPublish := publishArtifact
	t.Cleanup(func() { publishArtifact = originalPublish })
	publishArtifact = func(string, []byte, func(string) error) (coordination.PublishResult, error) {
		return coordination.PublishResult{}, errors.New("injected publication failure")
	}
	if err := publish(plan{files: []plannedFile{{path: target, body: []byte("new")}}}); !errors.Is(err, ErrPublish) {
		t.Fatalf("error = %v", err)
	}
	if _, err := os.Stat(filepath.Join(root, "src")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("new source directories remain after rollback: %v", err)
	}
}

func TestPublishRollbackRestoresRemovedStaleFilesAndModes(t *testing.T) {
	root := t.TempDir()
	updated := filepath.Join(root, "Updated.bas")
	staleFirst := filepath.Join(root, "StaleFirst.bas")
	staleSecond := filepath.Join(root, "StaleSecond.bas")
	for path, body := range map[string]string{updated: "old updated", staleFirst: "old first", staleSecond: "old second"} {
		writeTestFile(t, path, body)
	}
	if err := os.Chmod(updated, 0o600); err != nil {
		t.Fatal(err)
	}
	originalRemove := removeArtifact
	t.Cleanup(func() { removeArtifact = originalRemove })
	removeArtifact = func(path string) error {
		if path == staleSecond {
			return errors.New("injected stale removal failure")
		}
		return originalRemove(path)
	}
	p := plan{
		files: []plannedFile{{path: updated, body: []byte("new updated")}},
		stale: []string{staleFirst, staleSecond},
	}
	if err := publish(p); !errors.Is(err, ErrPublish) {
		t.Fatalf("error = %v", err)
	}
	for path, want := range map[string]string{updated: "old updated", staleFirst: "old first", staleSecond: "old second"} {
		body, err := os.ReadFile(path)
		if err != nil || string(body) != want {
			t.Fatalf("rollback %s = %q, %v", path, body, err)
		}
	}
	info, err := os.Stat(updated)
	if err != nil {
		t.Fatal(err)
	}
	if got := info.Mode().Perm(); runtime.GOOS != "windows" && got != 0o600 {
		t.Fatalf("restored mode = %v, want 0600", got)
	}
}

func TestRemoveGeneratedLineNumbers(t *testing.T) {
	source := "Public Sub Run()\n2  Debug.Print \"ok\"\nEnd Sub\n"
	got, err := removeGeneratedLineNumbers(source)
	if err != nil {
		t.Fatal(err)
	}
	if got != "Public Sub Run()\nDebug.Print \"ok\"\nEnd Sub\n" {
		t.Fatalf("got %q", got)
	}
	if _, err := removeGeneratedLineNumbers("Public Sub Run()\n10  Debug.Print \"bad\"\nEnd Sub\n"); err == nil {
		t.Fatal("expected mismatched physical line error")
	}
}

func testConfig() config.Config {
	cfg := config.Default()
	cfg.Excel.Path = "Book.xlsm"
	return cfg
}

func readFixture(t testing.TB, name string) []byte {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "pack", "vbaproject", "testdata", "corpus", name))
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func readPictureFixture(t testing.TB, name string) []byte {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", name))
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func excelAuthoredPictureAssets(t testing.TB) map[string]picture.Asset {
	t.Helper()
	project, err := vbaproject.Read(readPictureFixture(t, "baseline.bin"))
	if err != nil {
		t.Fatal(err)
	}
	assets := make(map[string]picture.Asset)
	var visit func([]*oforms.Control)
	visit = func(controls []*oforms.Control) {
		for _, control := range controls {
			if control == nil {
				continue
			}
			if control.Kind == "MSForms.Image" && control.Record != nil {
				if raw, ok := control.Record.Pictures["Picture"]; ok {
					prefix := []byte{
						0x04, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11,
						0x9d, 0xe3, 0x00, 0xaa, 0x00, 0x4b, 0xb8, 0x51,
						0x6c, 0x74, 0x00, 0x00, 0xb6, 0x04, 0x00, 0x00, 0x42, 0x4d,
					}
					if !bytes.HasPrefix(raw, prefix) {
						t.Fatalf("%s picture persistence preamble = % x, want % x", control.Name, raw[:min(len(raw), len(prefix))], prefix)
					}
					asset, err := picture.Decode(raw)
					if err != nil {
						t.Fatalf("decode Excel-authored picture for %s: %v", control.Name, err)
					}
					assets[control.Name] = asset
				}
			}
			children := control.Children
			if control.MultiPage != nil {
				children = control.MultiPage.Pages
			}
			visit(children)
		}
	}
	if len(project.Forms) != 1 {
		t.Fatalf("Excel-authored fixture form count = %d, want 1", len(project.Forms))
	}
	visit(project.Forms[0].Controls)
	if len(assets) != 2 {
		t.Fatalf("Excel-authored fixture image count = %d, want 2", len(assets))
	}
	return assets
}

func containsPath(paths []string, path string) bool {
	return slices.Contains(paths, path)
}

func writeWorkbook(t testing.TB, root string, project []byte) string {
	t.Helper()
	path := filepath.Join(root, "Book.xlsm")
	writeZip(t, path, project)
	return path
}

func writeZip(t testing.TB, path string, project []byte) {
	t.Helper()
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	if project != nil {
		entry, err := w.Create("xl/vbaProject.bin")
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write(project); err != nil {
			t.Fatal(err)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, path, buf.String())
}

func writeTestFile(t testing.TB, path, body string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func replaceCFBStream(t testing.TB, data []byte, target string, replace func([]byte) []byte) []byte {
	t.Helper()
	container, err := cfb.Open(data)
	if err != nil {
		t.Fatal(err)
	}
	writer, err := cfb.NewWriterForFormat(container.Format())
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range container.StoragePaths() {
		meta, _ := container.Storage(path)
		var parts []string
		if path != "" {
			parts = strings.Split(path, "/")
		}
		writer.AddStorage(parts, meta)
	}
	for _, path := range container.Paths() {
		body, _ := container.Stream(path)
		if path == target {
			body = replace(body)
		}
		writer.AddStream(strings.Split(path, "/"), body)
	}
	result, err := writer.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	return result
}
