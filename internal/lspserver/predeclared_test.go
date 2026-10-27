package lspserver

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func TestWorkspaceVB029PredeclaredClassOverlay(t *testing.T) {
	root := t.TempDir()
	classPath := filepath.Join(root, "src", "classes", "Predeclared.cls")
	mainPath := filepath.Join(root, "src", "modules", "Main.bas")
	classSource := "Attribute VB_Name = \"Predeclared\"\nAttribute VB_PredeclaredId = False\nOption Explicit\n"
	mainSource := "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n  Set Predeclared = Nothing\nEnd Sub\n"
	for path, source := range map[string]string{classPath: classSource, mainPath: mainSource} {
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	server, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := server.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}
	doc := intel.Document{URI: pathToFileURI(mainPath), Path: mainPath, Source: mainSource, ModuleKind: "standard", Version: 1}
	classDoc := intel.Document{URI: pathToFileURI(classPath), Path: classPath, ModuleKind: "class"}
	for version, value := range []string{"False", "True", "False"} {
		classDoc.Source = strings.Replace(classSource, "False", value, 1)
		classDoc.Version = int32(version + 1)
		// Supply the current open class without saving it to disk. The current
		// document remains unchanged while dependency attribute edits take effect.
		analyzer := server.analyzer
		analyzer.WorkspaceSymbolsSnapshotFunc = func(open []intel.Document) ([]intel.Symbol, error) {
			return server.cachedWorkspaceSymbolsSnapshot(append(open, classDoc))
		}
		count := 0
		for _, diagnostic := range analyzer.DiagnosticsContext(t.Context(), doc) {
			if diagnostic.Code == "VB029" {
				count++
			}
		}
		want := 1
		if value == "True" {
			want = 0
		}
		if count != want {
			t.Fatalf("version %d, VB_PredeclaredId=%s: VB029 count = %d, want %d", classDoc.Version, value, count, want)
		}
	}
}
