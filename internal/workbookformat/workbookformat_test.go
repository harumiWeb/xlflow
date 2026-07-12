package workbookformat

import (
	"errors"
	"testing"
)

func TestValidateFormulaSnapshotWorkbookUnsupportedProjectFormats(t *testing.T) {
	for _, tc := range []struct {
		path string
		want string
	}{
		{path: "Addin.xlam", want: ExtXLAM},
		{path: "Model.xlsb", want: ExtXLSB},
	} {
		t.Run(tc.path, func(t *testing.T) {
			err := ValidateFormulaSnapshotWorkbook(tc.path)
			var unsupported UnsupportedError
			if !errors.As(err, &unsupported) {
				t.Fatalf("error = %v, want UnsupportedError", err)
			}
			if unsupported.Capability != "formulas pull" || unsupported.Extension != tc.want {
				t.Fatalf("unsupported = %+v, want capability formulas pull extension %s", unsupported, tc.want)
			}
		})
	}
}

func TestValidateFileInspectWorkbookUnsupportedProjectFormats(t *testing.T) {
	for _, tc := range []struct {
		path string
		want string
	}{
		{path: "Addin.xlam", want: ExtXLAM},
		{path: "Model.xlsb", want: ExtXLSB},
	} {
		t.Run(tc.path, func(t *testing.T) {
			err := ValidateFileInspectWorkbook(tc.path, "inspect workbook")
			var unsupported UnsupportedError
			if !errors.As(err, &unsupported) {
				t.Fatalf("error = %v, want UnsupportedError", err)
			}
			if unsupported.Capability != "inspect workbook" || unsupported.Extension != tc.want {
				t.Fatalf("unsupported = %+v, want capability inspect workbook extension %s", unsupported, tc.want)
			}
		})
	}
}
