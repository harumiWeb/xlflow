package lint

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestBatchIRReuseMatchesFreshCanonicalBuild(t *testing.T) {
	root := t.TempDir()
	for _, name := range []string{"Main.bas", "Widget.cls", "Dialog.frm"} {
		t.Run(name, func(t *testing.T) {
			path := filepath.Join(root, "src", "modules", name)
			source := []byte("Attribute VB_Name = \"Renamed\"\nOption Explicit\nPublic Sub Main()\nDim item As Widget\nCall Missing(1)\nEnd Sub\n")
			doc, err := vbaast.ParseDocument(path, source)
			if err != nil {
				t.Fatal(err)
			}
			original, err := procedureir.BuildParsedContext(t.Context(), procedureir.BuildOptions{Path: path}, doc)
			doc.Close()
			if err != nil {
				t.Fatal(err)
			}
			linter := Linter{RootDir: root, batchIRCache: make(map[string]batchIRCacheEntry)}
			linter.rememberBatchIR(path, source, original)
			got, reused := linter.reuseBatchIR(path, source, original.ModuleName, original.ModuleKind)
			want, err := procedureir.BuildSource(procedureir.BuildOptions{
				RootDir: root, Path: path, ModuleName: original.ModuleName, ModuleKind: original.ModuleKind,
			}, source)
			// The canonical clone normalizes some empty slices to nil. Compare
			// owned snapshots; CLI diagnostic equality is checked separately.
			want = procedureir.CloneDocumentIR(want)
			if err != nil || !reused || !reflect.DeepEqual(got, want) {
				t.Fatalf("canonical IR differs: reused=%v error=%v", reused, err)
			}
			got.Procedures[0].Calls[0].Callee.Text = "mutated"
			again, reused := linter.reuseBatchIR(path, source, original.ModuleName, original.ModuleKind)
			if !reused || !reflect.DeepEqual(again, want) {
				t.Fatal("consumer mutated cached IR")
			}
			for _, mismatch := range []struct{ source, module, kind string }{
				{string(source) + "' changed", original.ModuleName, original.ModuleKind},
				{string(source), "Other", original.ModuleKind},
				{string(source), original.ModuleName, "document"},
				{string(source), "", original.ModuleKind},
			} {
				if _, ok := linter.reuseBatchIR(path, []byte(mismatch.source), mismatch.module, mismatch.kind); ok {
					t.Fatal("reused incompatible source or module identity")
				}
			}
		})
	}
}

func TestBatchIRReuseNeverStoresIncompleteParses(t *testing.T) {
	linter := Linter{batchIRCache: make(map[string]batchIRCacheEntry)}
	for _, summary := range []procedureir.ParseSummary{{HasError: true}, {HasMissing: true}} {
		linter.rememberBatchIR("Main.bas", nil, procedureir.DocumentIR{Parse: summary})
	}
	if len(linter.batchIRCache) != 0 {
		t.Fatal("incomplete parse was cached")
	}
}

func TestBatchIRReuseIsDiscardedAcrossRunsAndFilteredCalls(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o700); err != nil {
		t.Fatal(err)
	}
	write := func(body string) {
		t.Helper()
		if err := os.WriteFile(path, []byte("Attribute VB_Name = \"Main\"\nOption Explicit\n"+body), 0o600); err != nil {
			t.Fatal(err)
		}
	}
	stale := map[string]batchIRCacheEntry{"sentinel": {}}
	linter := Linter{RootDir: root, Config: config.Default(), batchIRCache: stale}
	write("Public Sub Main()\nEnd Sub\n")
	if _, err := linter.RunResultContext(t.Context()); err != nil {
		t.Fatal(err)
	}
	write("Public Sub Main()\nMissing = 1\nEnd Sub\n")
	result, err := linter.RunResultContext(t.Context())
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, issue := range result.Issues {
		found = found || issue.Code == "VB029"
	}
	if !found || len(stale) != 1 {
		t.Fatal("run reused stale evidence or retained its cache")
	}
	linter.PathFilter = func(string) bool { return true }
	if _, err := linter.RunResultContext(t.Context()); err != nil || len(stale) != 1 {
		t.Fatalf("filtered call retained cache: %v", err)
	}
	ctx, cancel := context.WithCancel(t.Context())
	cancel()
	if _, err := linter.RunResultContext(ctx); !errors.Is(err, context.Canceled) {
		t.Fatalf("cancellation was lost: %v", err)
	}
}

func TestBatchIRReusePreservesProjectDiagnostics(t *testing.T) {
	for _, tc := range []struct {
		name, path, source string
		wantCode           string
	}{
		{"resolution errors", "classes/Main.cls", "Public Event Changed(ByVal value As Long)\nPrivate Const Scalar = 1\nSub Run()\nScalar()\nRaiseEvent Missing(1)\nEnd Sub\n", "VB052"},
		{"conditional declarations", "modules/Main.bas", "#If VBA7 Then\nPublic Sub Main()\nEnd Sub\n#Else\nPublic Sub Main()\nEnd Sub\n#End If\n", ""},
		{"incomplete syntax", "modules/Main.bas", "Sub Main()\nIf Then\nMissing(\n", ""},
		{"worksheet event", "worksheets/Sheet1.cls", "Private Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n", ""},
	} {
		t.Run(tc.name, func(t *testing.T) {
			root := t.TempDir()
			path := filepath.Join(root, "src", filepath.FromSlash(tc.path))
			if err := os.MkdirAll(filepath.Dir(path), 0o700); err != nil {
				t.Fatal(err)
			}
			if err := os.WriteFile(path, []byte(tc.source), 0o600); err != nil {
				t.Fatal(err)
			}
			linter := Linter{RootDir: root, Config: config.Default(), batchIRCache: make(map[string]batchIRCacheEntry)}
			if _, err := linter.lintFileContext(t.Context(), path); err != nil {
				t.Fatal(err)
			}
			got, err := linter.projectIssuesContext(t.Context(), []string{path})
			if err != nil {
				t.Fatal(err)
			}
			linter.batchIRCache = nil
			want, err := linter.projectIssuesContext(t.Context(), []string{path})
			if err != nil || !reflect.DeepEqual(got, want) {
				t.Fatalf("project diagnostics changed: got=%+v want=%+v err=%v", got, want, err)
			}
			if tc.wantCode != "" && len(issuesByCode(got, tc.wantCode)) == 0 {
				t.Fatalf("negative control lost %s: %+v", tc.wantCode, got)
			}
		})
	}
}
