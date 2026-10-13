package cli

import (
	"os"
	"path/filepath"
	"strconv"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
)

func TestScaffoldCodePageUsesWorkbookEncoding(t *testing.T) {
	for _, codePage := range []uint16{932, 1252} {
		t.Run(strconv.Itoa(int(codePage)), func(t *testing.T) {
			root := t.TempDir()
			cfg := config.Default()
			cfg.Excel.Path = "Book.xlsm"
			if err := config.Write(filepath.Join(root, config.FileName), cfg); err != nil {
				t.Fatal(err)
			}
			body, _, err := packpkg.BuildBlankWorkbook([]packpkg.SourceModule{
				{Name: "ThisWorkbook", Type: packpkg.ModuleTypeDocument},
				{Name: "Sheet1", Type: packpkg.ModuleTypeDocument},
			}, packpkg.BlankOptions{CodePage: codePage})
			if err != nil {
				t.Fatal(err)
			}
			if err := os.WriteFile(filepath.Join(root, cfg.Excel.Path), body, 0o600); err != nil {
				t.Fatal(err)
			}
			if err := seedScaffoldCodePage(root); err != nil {
				t.Fatal(err)
			}
			got, err := config.Load(root)
			if err != nil {
				t.Fatal(err)
			}
			if got.Pack.Blank.CodePage != codePage {
				t.Fatalf("blank codepage = %d, workbook codepage = %d", got.Pack.Blank.CodePage, codePage)
			}
		})
	}
}
