package excel

import (
	"archive/zip"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
)

func TestRunnerInspectFormProjectsPersistedDisplayedSizeIntoBothDesignerSnapshot(t *testing.T) {
	projectBytes, expectedWidth, expectedHeight := userFormSnapshotFixture(t)
	var request excelbridge.Request
	var callCount int
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(root string, mode excelbridge.Mode) excelbridge.Provider {
		return callbackBridgeProvider{
			name: string(mode),
			execute: func(req excelbridge.Request) {
				callCount++
				request = req
				if err := writeWorkbookSnapshot(req.Args["PrivateSnapshotWorkbookPath"], projectBytes); err != nil {
					t.Errorf("write private workbook snapshot: %v", err)
				}
			},
			responseFunc: func(excelbridge.Request) excelbridge.Response {
				body, err := json.Marshal(map[string]any{
					"protocol_version": excelbridge.ProtocolVersion,
					"status":           output.StatusOK,
					"command":          "inspect-form",
					"forms": map[string]any{
						"runtime": map[string]any{
							"name": "UserForm1", "basis": "runtime", "client_width": 111.0, "client_height": 112.0,
						},
						"designer": map[string]any{
							"name": "UserForm1", "basis": "designer", "width": 530.0, "height": 430.0,
							"client_width": 520.2, "client_height": 420.00005,
						},
					},
				})
				if err != nil {
					t.Errorf("marshal inspect response: %v", err)
				}
				return excelbridge.Response{Stdout: body}
			},
		}
	}

	env, code, err := (Runner{RootDir: t.TempDir(), BridgeMode: "dotnet"}).InspectForm(config.Default(), InspectFormOptions{
		Name:    "UserForm1",
		Basis:   "both",
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Status != output.StatusOK {
		t.Fatalf("inspect result = status %q, exit %d, error %+v", env.Status, code, env.Error)
	}
	if callCount != 1 {
		t.Fatalf("bridge calls = %d, want one inspection call", callCount)
	}
	snapshotPath := request.Args["PrivateSnapshotWorkbookPath"]
	if !filepath.IsAbs(snapshotPath) || filepath.Ext(snapshotPath) != ".xlsm" {
		t.Fatalf("private snapshot path = %q, want absolute .xlsm path", snapshotPath)
	}
	if _, err := os.Stat(filepath.Dir(snapshotPath)); !os.IsNotExist(err) {
		t.Fatalf("private snapshot directory was not removed: stat error = %v", err)
	}

	forms, ok := env.Forms.(map[string]any)
	if !ok {
		t.Fatalf("forms payload type = %T", env.Forms)
	}
	runtimeForm := forms["runtime"].(map[string]any)
	if runtimeForm["client_width"] != 111.0 || runtimeForm["client_height"] != 112.0 {
		t.Fatalf("runtime snapshot was modified: %+v", runtimeForm)
	}
	designerForm := forms["designer"].(map[string]any)
	if designerForm["client_width"] != expectedWidth || designerForm["client_height"] != expectedHeight {
		t.Fatalf("designer client dimensions = %v x %v, want DisplayedSize projection %v x %v", designerForm["client_width"], designerForm["client_height"], expectedWidth, expectedHeight)
	}
	if designerForm["observed_client_width"] != 520.2 || designerForm["observed_client_height"] != 420.00005 {
		t.Fatalf("exported VBFrame dimensions were not retained as observations: %+v", designerForm)
	}
}

func TestProjectPersistedFormDimensionsUpdatesSingleDesignerSnapshot(t *testing.T) {
	projectBytes, expectedWidth, expectedHeight := userFormSnapshotFixture(t)
	snapshotPath := filepath.Join(t.TempDir(), "snapshot.xlsm")
	if err := writeWorkbookSnapshot(snapshotPath, projectBytes); err != nil {
		t.Fatal(err)
	}
	var formsPayload any = map[string]any{
		"name": "UserForm1", "basis": "designer", "client_width": 520.2, "client_height": 420.00005,
	}
	if err := projectPersistedFormDimensions(snapshotPath, "UserForm1", "designer", &formsPayload); err != nil {
		t.Fatal(err)
	}
	designer := formsPayload.(map[string]any)
	if designer["client_width"] != expectedWidth || designer["client_height"] != expectedHeight {
		t.Fatalf("single designer client dimensions = %v x %v, want %v x %v", designer["client_width"], designer["client_height"], expectedWidth, expectedHeight)
	}
}

func TestReadRootDisplayedSizeResolvesStorageNameCaseInsensitively(t *testing.T) {
	projectBytes, expectedWidth, expectedHeight := userFormSnapshotFixture(t)
	snapshotPath := filepath.Join(t.TempDir(), "snapshot.xlsm")
	if err := writeWorkbookSnapshot(snapshotPath, projectBytes); err != nil {
		t.Fatal(err)
	}

	got, err := readRootDisplayedSize(snapshotPath, "userform1")
	if err != nil {
		t.Fatal(err)
	}
	if width, height := float64(got.Width)*72/2540, float64(got.Height)*72/2540; width != expectedWidth || height != expectedHeight {
		t.Fatalf("lower-case form lookup dimensions = %v x %v, want %v x %v", width, height, expectedWidth, expectedHeight)
	}
}

func TestProjectPersistedFormDimensionsSkipsRuntimeSnapshot(t *testing.T) {
	var formsPayload any = map[string]any{
		"name": "UserForm1", "basis": "runtime", "client_width": 111.0, "client_height": 112.0,
	}
	if err := projectPersistedFormDimensions("missing.xlsm", "UserForm1", "runtime", &formsPayload); err != nil {
		t.Fatal(err)
	}
	runtimeForm := formsPayload.(map[string]any)
	if runtimeForm["client_width"] != 111.0 || runtimeForm["client_height"] != 112.0 {
		t.Fatalf("runtime snapshot changed: %+v", formsPayload)
	}
}

func userFormSnapshotFixture(t *testing.T) ([]byte, float64, float64) {
	t.Helper()
	projectBytes, err := os.ReadFile(filepath.Join("..", "pack", "vbaproject", "testdata", "corpus", "p4_form.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(projectBytes)
	if err != nil {
		t.Fatal(err)
	}
	for _, form := range project.Forms {
		if form.Name != "UserForm1" || len(form.Levels) == 0 || form.Levels[0].Record == nil {
			continue
		}
		size, ok := form.Levels[0].Record.Sizes["DisplayedSize"]
		if !ok {
			t.Fatal("fixture UserForm1 is missing root DisplayedSize")
		}
		return projectBytes, float64(size.Width) * 72 / 2540, float64(size.Height) * 72 / 2540
	}
	t.Fatal("fixture does not contain UserForm1")
	return nil, 0, 0
}

func writeWorkbookSnapshot(path string, projectBytes []byte) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	writer := zip.NewWriter(file)
	entry, err := writer.Create("xl/vbaProject.bin")
	if err == nil {
		_, err = entry.Write(projectBytes)
	}
	closeErr := writer.Close()
	fileErr := file.Close()
	if err != nil {
		return err
	}
	if closeErr != nil {
		return closeErr
	}
	return fileErr
}
