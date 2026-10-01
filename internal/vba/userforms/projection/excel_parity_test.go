package projection

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const (
	parityWorkbookEnv = "XLFLOW_PROJECTION_PARITY_WORKBOOK"
	paritySnapshotEnv = "XLFLOW_PROJECTION_PARITY_SNAPSHOT"
)

// TestProjectMatchesExcelBackedSnapshot compares both readers against the
// same workbook. Set the two parity environment variables to run this
// Windows/Excel integration check; ordinary pure-Go test runs skip it.
func TestProjectMatchesExcelBackedSnapshot(t *testing.T) {
	workbookPath := os.Getenv(parityWorkbookEnv)
	snapshotPath := os.Getenv(paritySnapshotEnv)
	if workbookPath == "" || snapshotPath == "" {
		t.Skipf("set %s and %s to run Excel snapshot parity", parityWorkbookEnv, paritySnapshotEnv)
	}

	format := strings.TrimPrefix(filepath.Ext(snapshotPath), ".")
	excelSnapshot, err := forms.LoadFormSpec(forms.SpecInput{
		Path: snapshotPath, DisplayPath: snapshotPath, Format: format,
	})
	if err != nil {
		t.Fatalf("load Excel-backed snapshot: %v", err)
	}

	projectBytes, err := workbookVBAProject(workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(projectBytes)
	if err != nil {
		t.Fatalf("read workbook VBA project: %v", err)
	}
	container, err := cfb.Open(projectBytes)
	if err != nil {
		t.Fatalf("open workbook VBA project: %v", err)
	}
	form, err := oforms.ReadForm(container, excelSnapshot.Form.Name, project.Props.CodePage)
	if err != nil {
		t.Fatalf("read MS-OFORMS for %q: %v", excelSnapshot.Form.Name, err)
	}
	pureSnapshot, err := Project(form)
	if err != nil {
		t.Fatalf("project MS-OFORMS: %v", err)
	}
	if err := compareSupportedSnapshotState(excelSnapshot, pureSnapshot); err != nil {
		t.Fatal(err)
	}
}

func compareSupportedSnapshotState(excel, pure forms.FormSpec) error {
	if excel.Form.Name != pure.Form.Name || optionalString(excel.Form.Caption) != optionalString(pure.Form.Caption) {
		return fmt.Errorf("form identity differs: Excel=%#v pure-Go=%#v", excel.Form, pure.Form)
	}
	// FormSpec documents root size as best-effort: the binary stores client
	// dimensions while Excel reports its outer Designer dimensions.
	if len(excel.Controls) != len(pure.Controls) {
		return fmt.Errorf("control count differs: Excel=%d pure-Go=%d", len(excel.Controls), len(pure.Controls))
	}
	for index, excelControl := range excel.Controls {
		pureControl := pure.Controls[index]
		if excelControl.ID != pureControl.ID || excelControl.ParentID != pureControl.ParentID ||
			excelControl.Name != pureControl.Name || excelControl.Type != pureControl.Type ||
			(excelControl.ProgID != "" && pureControl.ProgID != "" && !strings.EqualFold(excelControl.ProgID, pureControl.ProgID)) ||
			optionalInt(excelControl.ZIndex) != optionalInt(pureControl.ZIndex) {
			return fmt.Errorf("control %d identity differs: Excel=%#v pure-Go=%#v", index, excelControl, pureControl)
		}
		if optionalString(excelControl.Caption) != optionalString(pureControl.Caption) ||
			optionalString(excelControl.Text) != optionalString(pureControl.Text) ||
			!sameJSONValue(excelControl.Value, pureControl.Value) ||
			!slices.Equal(excelControl.List, pureControl.List) {
			return fmt.Errorf("control %q content differs: Excel=%#v pure-Go=%#v", excelControl.Name, excelControl, pureControl)
		}
		if !nearFloat(excelControl.Left, pureControl.Left) || !nearFloat(excelControl.Top, pureControl.Top) ||
			!nearFloat(excelControl.Width, pureControl.Width) || !nearFloat(excelControl.Height, pureControl.Height) {
			return fmt.Errorf("control %q geometry differs: Excel=%#v pure-Go=%#v", excelControl.Name, excelControl, pureControl)
		}
		if optionalInt(excelControl.TabIndex) != optionalInt(pureControl.TabIndex) ||
			optionalInt(excelControl.SelectedIndex) != optionalInt(pureControl.SelectedIndex) ||
			optionalBool(excelControl.Enabled) != optionalBool(pureControl.Enabled) ||
			optionalBool(excelControl.Visible) != optionalBool(pureControl.Visible) {
			return fmt.Errorf("control %q state differs: Excel=%#v pure-Go=%#v", excelControl.Name, excelControl, pureControl)
		}
	}
	return nil
}

func workbookVBAProject(workbookPath string) ([]byte, error) {
	archive, err := zip.OpenReader(workbookPath)
	if err != nil {
		return nil, err
	}
	defer func() { _ = archive.Close() }()
	for _, file := range archive.File {
		if file.Name != "xl/vbaProject.bin" {
			continue
		}
		stream, err := file.Open()
		if err != nil {
			return nil, err
		}
		body, readErr := io.ReadAll(stream)
		closeErr := stream.Close()
		if readErr != nil {
			return nil, readErr
		}
		if closeErr != nil {
			return nil, closeErr
		}
		return body, nil
	}
	return nil, fmt.Errorf("%s does not contain xl/vbaProject.bin", workbookPath)
}

func sameJSONValue(a, b any) bool {
	first, firstErr := json.Marshal(a)
	second, secondErr := json.Marshal(b)
	return firstErr == nil && secondErr == nil && bytes.Equal(first, second)
}

func nearFloat(a, b *float64) bool {
	if a == nil || b == nil {
		return a == nil && b == nil
	}
	return math.Abs(*a-*b) <= 0.05
}

func optionalString(value *string) string {
	if value == nil {
		return ""
	}
	return *value
}

func optionalInt(value *int) int {
	if value == nil {
		return -1
	}
	return *value
}

func optionalBool(value *bool) bool {
	return value != nil && *value
}
