package cli

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/picture"
)

func TestCanonicalPictureSourcePreflights(t *testing.T) {
	for _, command := range []string{"pack", "push"} {
		for _, assetPath := range []string{"src/modules/logo.bas", "src/classes/logo.cls", "src/workbook/logo.bas", "src/modules/logo.bmp", "src/forms/images/logo.bas", "src/forms/images/logo.cls", "src/forms/images/logo.frm", "src/forms/code/Login.bas", "src/forms/Linked.bas"} {
			t.Run(command+"/"+assetPath, func(t *testing.T) {
				dir := t.TempDir()
				writePackConfig(t, dir)
				writePushSourceTree(t, dir)
				writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
				stubClosedWorkbook(t)
				bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
				if err != nil {
					t.Fatal(err)
				}
				asset := filepath.Join(dir, filepath.FromSlash(assetPath))
				if assetPath == "src/forms/Linked.bas" {
					writePackSourceModule(t, dir, "assets/picture.bmp", bmp)
					if err := os.Symlink(filepath.Join("..", "..", "assets", "picture.bmp"), asset); err != nil {
						t.Skipf("cannot create picture symlink: %v", err)
					}
				} else {
					writePackSourceModule(t, dir, assetPath, bmp)
				}
				writePackSourceModule(t, dir, "src/forms/specs/Login.json", []byte(fmt.Sprintf(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":%q}}]}`, assetPath)))
				args := []string{"--json", "pack", "--out", "dist/Book.xlsm"}
				outPath := filepath.Join(dir, "dist", "Book.xlsm")
				if command == "push" {
					args = []string{"--json", "push", "--backend", "file"}
					outPath = filepath.Join(dir, "build", "Book.xlsm")
				}
				stdout, err := runBuildCommandForTest(dir, args...)
				if err != nil {
					t.Fatalf("canonical picture %s: %v\n%s", command, err, stdout)
				}
				project := readPackedProjectForTest(t, outPath)
				for _, module := range project.Modules {
					if module.Name == "logo" || module.Name == "Linked" {
						t.Fatalf("picture imported as module: %s", module.Name)
					}
				}
				if len(project.Forms) != 1 {
					t.Fatalf("forms = %d, want 1", len(project.Forms))
				}
				// Read back the persisted picture, not just the successful exit.
				decoded, err := picture.Decode(project.Forms[0].Controls[0].Record.Pictures["Picture"])
				if err != nil || !bytes.Equal(decoded.Data, bmp) {
					t.Fatalf("persisted picture bytes missing: %v", err)
				}
				before := readFileForTest(t, outPath)
				writePackSourceModule(t, dir, "src/modules/Bad.bas", []byte{0x81})
				stdout, err = runBuildCommandForTest(dir, args...)
				if err == nil || output.ExitCode(err) != output.ExitValidation || jsonErrorCode(t, stdout) != "source_encoding_invalid" {
					t.Fatalf("invalid VBA must still fail: %v\n%s", err, stdout)
				}
				if !bytes.Equal(before, readFileForTest(t, outPath)) {
					t.Fatal("invalid VBA changed workbook")
				}
			})
		}
	}
}

func TestCanonicalPictureCaseDistinctSidecar(t *testing.T) {
	for _, command := range []string{"pack", "push"} {
		t.Run(command, func(t *testing.T) {
			dir := t.TempDir()
			writePackConfig(t, dir)
			writePushSourceTree(t, dir)
			writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
			bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
			if err != nil {
				t.Fatal(err)
			}
			writePackSourceModule(t, dir, "src/forms/code/login.bas", bmp)
			if _, err := os.Stat(filepath.Join(dir, "src", "forms", "code", "Login.bas")); err == nil {
				t.Skip("requires a case-sensitive filesystem")
			} else if !os.IsNotExist(err) {
				t.Fatal(err)
			}
			writePackSourceModule(t, dir, "src/forms/code/Login.bas", []byte("Private Sub KeepMe()\n    Debug.Print \"case-sensitive\"\nEnd Sub\n"))
			writePackSourceModule(t, dir, "src/forms/specs/Login.json", []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/code/login.bas"}}]}`))
			args := []string{"--json", "pack", "--out", "dist/Book.xlsm"}
			outPath := filepath.Join(dir, "dist", "Book.xlsm")
			if command == "push" {
				args = []string{"--json", "push", "--backend", "file"}
				outPath = filepath.Join(dir, "build", "Book.xlsm")
			}
			stdout, err := runBuildCommandForTest(dir, args...)
			if err != nil {
				t.Fatalf("case-distinct picture %s: %v\n%s", command, err, stdout)
			}
			project := readPackedProjectForTest(t, outPath)
			if code := moduleSourceForTest(project, "Login"); !strings.Contains(code, "KeepMe") {
				t.Fatalf("case-distinct sidecar lost: %q", code)
			}
			decoded, err := picture.Decode(project.Forms[0].Controls[0].Record.Pictures["Picture"])
			if err != nil || !bytes.Equal(decoded.Data, bmp) {
				t.Fatalf("case-distinct picture lost: %v", err)
			}
			before := readFileForTest(t, outPath)
			writePackSourceModule(t, dir, "src/forms/code/Login.bas", []byte{0x81})
			stdout, err = runBuildCommandForTest(dir, args...)
			if err == nil || jsonErrorCode(t, stdout) != "source_encoding_invalid" {
				t.Fatalf("case-distinct sidecar bypassed encoding: %v\n%s", err, stdout)
			}
			if !bytes.Equal(before, readFileForTest(t, outPath)) {
				t.Fatal("invalid sidecar changed workbook")
			}
		})
	}
}
