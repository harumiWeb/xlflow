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

func TestContractImplicitApplicationCallbacksUseResolvedSnapshot(t *testing.T) {
	for _, test := range []struct {
		name, body, declarations, api string
		shadow                        bool
	}{
		{name: "receiverless Run", body: `Run "Worker.Target"`, api: "application.run"},
		{name: "receiverless OnTime", body: `OnTime Now, "Worker.Target"`, api: "application.ontime"},
		{name: "receiverless OnKey", body: `OnKey "{LEFT}", "Worker.Target"`, api: "application.onkey"},
		{name: "With Run", body: "With Application\n.Run \"Worker.Target\"\nEnd With", api: "application.run"},
		{name: "With OnTime", body: "With Application\n.OnTime Now, \"Worker.Target\"\nEnd With", api: "application.ontime"},
		{name: "With OnKey", body: "With Application\n.OnKey \"{LEFT}\", \"Worker.Target\"\nEnd With", api: "application.onkey"},
		{name: "project procedure shadow", body: `Run "Worker.Target"`, declarations: "Private Sub Run(ByVal target As String)\nEnd Sub\n", shadow: true},
		{name: "non-callable local shadow", body: "Dim Run As Long\nRun \"Worker.Target\"", shadow: true},
	} {
		t.Run(test.name, func(t *testing.T) {
			root, cfg := contractProject(t, 0)
			cfg.Project.Entry = "Main.Start"
			writeCollectorSource(t, root, "src/modules/Main.bas", "Attribute VB_Name = \"Main\"\nPublic Sub Start()\n"+test.body+"\nEnd Sub\n"+test.declarations)
			writeCollectorSource(t, root, "src/modules/Worker.bas", "Attribute VB_Name = \"Worker\"\nPrivate Sub Target()\nEnd Sub\n")
			report, _, err := collectContextWithHooks(t.Context(), root, cfg, Options{}, &collectionHooks{loadTypeDB: func() (typedb.LoadResult, error) {
				return typedb.LoadResult{DB: vbadb.New(), Complete: true}, nil
			}})
			if err != nil {
				t.Fatal(err)
			}
			var caller, target *Procedure
			for index := range report.Procedures {
				procedure := &report.Procedures[index]
				if procedure.QualifiedName == "Main.Start" {
					caller = procedure
				}
				if procedure.QualifiedName == "Worker.Target" {
					target = procedure
				}
			}
			if caller == nil || target == nil {
				t.Fatalf("missing procedures: %+v", report.Procedures)
			}
			if test.shadow {
				if len(report.DynamicReferences) != 0 || target.Reachability != "unreachable" {
					t.Fatalf("shadow invented callback: refs=%+v target=%+v", report.DynamicReferences, target)
				}
				return
			}
			if len(report.DynamicReferences) != 1 {
				t.Fatalf("missing callback: %+v", report.DynamicReferences)
			}
			ref := report.DynamicReferences[0]
			if ref.API != test.api || ref.Target != "Worker.Target" || ref.Kind != "static" || ref.CallerID != caller.ID || ref.Range.StartLine <= 0 {
				t.Fatalf("callback lost canonical evidence: %+v", ref)
			}
			if target.Reachability != "possible" || len(report.PossibleReachability) != 1 || report.PossibleReachability[0].ID.String() != target.CallgraphID {
				t.Fatalf("callback lost possible reachability: target=%+v possible=%+v", target, report.PossibleReachability)
			}
			for _, node := range report.Unreachable {
				if node.ID.String() == target.CallgraphID {
					t.Fatal("callback target was reported unreachable")
				}
			}
		})
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
