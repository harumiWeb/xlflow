package architecture

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func contractProject(t testing.TB, count int) (string, config.Config) {
	t.Helper()
	root := t.TempDir()
	cfg := config.Default()
	cfg.Project.Entry = "Main.Run"
	cfg.Src.Modules = "src/modules"
	if err := os.MkdirAll(filepath.Join(root, cfg.Src.Modules), 0o755); err != nil {
		t.Fatal(err)
	}
	for index := range count {
		name := fmt.Sprintf("Worker%03d", index)
		body := "Public Sub Work()\nEnd Sub\n"
		if index == 0 {
			name = "Main"
			body = "Private state As Long\nPublic Sub Run()\nstate = state + 1\nEnd Sub\nPrivate Sub Unused()\nEnd Sub\n"
		}
		if err := os.WriteFile(filepath.Join(root, cfg.Src.Modules, name+".bas"), []byte("Attribute VB_Name = \""+name+"\"\n"+body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	return root, cfg
}

func TestContractBuildsEachSourceOnceAndKeepsCapturedBytes(t *testing.T) {
	root, cfg := contractProject(t, 3)
	counts := map[string]int{}
	hooks := &collectionHooks{
		afterSourceRead: func(path string) {
			counts["read"]++
			// Root discovery must use captured source, not reopen this mutation.
			if err := os.WriteFile(filepath.Join(root, filepath.FromSlash(path)), []byte("broken syntax"), 0o644); err != nil {
				t.Fatal(err)
			}
		},
		afterParse:   func(string) { counts["parse"]++ },
		afterIRBuild: func(string) { counts["ir"]++ },
		afterCFG:     func(string) { counts["cfg"]++ },
	}
	report, _, err := collectContextWithHooks(t.Context(), root, cfg, Options{}, hooks)
	if err != nil {
		t.Fatal(err)
	}
	for _, kind := range []string{"read", "parse", "ir", "cfg"} {
		if counts[kind] != 3 {
			t.Errorf("%s construction count=%d, want3", kind, counts[kind])
		}
	}
	if report.Summary.Modules != 3 || report.Summary.Procedures != 4 || len(report.Unreachable) != 1 {
		t.Fatalf("captured snapshot lost facts: %+v", report.Summary)
	}
	ids := map[string]bool{}
	for _, procedure := range report.Procedures {
		ids[procedure.ID] = true
	}
	for _, field := range report.ModuleState.Fields {
		for _, accesses := range [][]string{field.Readers, field.Writers, field.Mutators} {
			for _, id := range accesses {
				if !ids[id] {
					t.Errorf("state access cannot join procedure %q", id)
				}
			}
		}
	}
}

func TestContractCancellationReturnsNoPartialSnapshot(t *testing.T) {
	root, cfg := contractProject(t, 2)
	ctx, cancel := context.WithCancel(t.Context())
	defer cancel()
	report, _, err := collectContextWithHooks(ctx, root, cfg, Options{}, &collectionHooks{afterSourceRead: func(string) { cancel() }})
	if !errors.Is(err, context.Canceled) || !reflect.DeepEqual(report, Report{}) {
		t.Fatalf("canceled snapshot: report=%+v err=%v", report.Summary, err)
	}
}

func TestContractPreservesCanonicalNonCallableShadowing(t *testing.T) {
	root, cfg := contractProject(t, 1)
	source := "Attribute VB_Name = \"Main\"\nPublic Sub Run()\nDim Work As Long\nDebug.Print Work(1)\nEnd Sub\nPublic Function Work(ByVal index As Long) As Long\nWork = index\nEnd Function\n"
	if err := os.WriteFile(filepath.Join(root, cfg.Src.Modules, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	report, _, err := collectContextWithHooks(t.Context(), root, cfg, Options{}, &collectionHooks{loadTypeDB: func() (typedb.LoadResult, error) {
		return typedb.LoadResult{DB: vbadb.New(), Complete: true}, nil
	}})
	if err != nil {
		t.Fatal(err)
	}
	if report.Summary.ConfirmedCallEdges != 0 {
		t.Fatalf("non-callable scalar invented a confirmed function call: %+v", report.Dependencies.Edges)
	}
	for _, evidence := range report.Uncertainty.Evidence {
		if evidence.Target == "Work" && evidence.Status == "non_callable" {
			return
		}
	}
	t.Fatalf("canonical non-callable evidence was lost: %+v", report.Uncertainty.Evidence)
}

func TestContractEmptyCollectionsAreArraysAndRepeatedReportsMatch(t *testing.T) {
	root, cfg := contractProject(t, 0)
	first, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	second, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	left, err := json.Marshal(first)
	if err != nil {
		t.Fatal(err)
	}
	right, err := json.Marshal(second)
	if err != nil {
		t.Fatal(err)
	}
	if string(left) != string(right) {
		t.Fatal("repeated snapshot JSON differs")
	}
	var value any
	if err := json.Unmarshal(left, &value); err != nil {
		t.Fatal(err)
	}
	var check func(any, string)
	check = func(node any, path string) {
		switch node := node.(type) {
		case map[string]any:
			for key, child := range node {
				if child == nil {
					t.Errorf("null collection at %s.%s", path, key)
				}
				check(child, path+"."+key)
			}
		case []any:
			for index, child := range node {
				check(child, fmt.Sprintf("%s[%d]", path, index))
			}
		}
	}
	check(value, "architecture")
}

func TestContractCollectsEmbeddedFormSource(t *testing.T) {
	root, cfg := contractProject(t, 1)
	cfg.UserForm.CodeSource = "frm"
	writeCollectorSource(t, root, "src/forms/Panel.frm", "VERSION 5.00\nBegin VB.UserForm Panel\n    Caption = \"Panel\"\n    ClientHeight = 1440\n    ClientWidth = 2880\nEnd\nAttribute VB_Name = \"Panel\"\nPrivate Sub UserForm_Initialize()\nEnd Sub\n")
	report, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if report.Summary.Modules != 2 || report.Summary.Procedures != 3 {
		t.Fatalf("embedded source was counted incorrectly: %+v", report.Summary)
	}
}

func BenchmarkArchitectureSparseProject(b *testing.B) {
	// Keep developer-specific generated TypeLib inventories out of this fixture.
	b.Setenv(typedb.EnvDir, b.TempDir())
	for _, size := range []int{16, 128} {
		b.Run(fmt.Sprint(size), func(b *testing.B) {
			root, cfg := contractProject(b, size)
			b.ReportAllocs()
			for b.Loop() {
				report, _, err := CollectContext(b.Context(), root, cfg, Options{})
				if err != nil || report.Summary.Modules != size {
					b.Fatalf("snapshot: modules=%d err=%v", report.Summary.Modules, err)
				}
			}
		})
	}
}
