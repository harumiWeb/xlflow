package build

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestPlanFiltersComponentsDeterministicallyWithoutMutatingInputs(t *testing.T) {
	root := writeSourceTree(t)
	cfg := config.Default()
	cfg.Build.Exclude = []string{"src\\modules\\Tests\\**", "src/forms/code/Login.bas", "src/missing/**"}
	before := readTree(t, root)

	plan, err := Plan(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if plan.BaseWorkbook != "build/Book.xlsm" || plan.OutputPath != "build/Release/Book.xlsm" {
		t.Fatalf("paths = %+v", plan)
	}
	if got, want := componentPaths(plan.Included), []string{"src/classes/Service.cls", "src/modules/Main.bas", "src/workbook/ThisWorkbook.bas"}; !reflect.DeepEqual(got, want) {
		t.Fatalf("included = %#v, want %#v", got, want)
	}
	if got, want := componentPaths(plan.Excluded), []string{"src/forms/Login.frm", "src/modules/Tests/TestMain.bas"}; !reflect.DeepEqual(got, want) {
		t.Fatalf("excluded = %#v, want %#v", got, want)
	}
	if got := plan.Excluded[0].RelatedPaths; !reflect.DeepEqual(got, []string{"src/forms/Login.frx", "src/forms/code/Login.bas"}) {
		t.Fatalf("form related paths = %#v", got)
	}
	if len(plan.Warnings) != 1 || plan.Warnings[0].Pattern != "src/missing/**" {
		t.Fatalf("warnings = %#v", plan.Warnings)
	}
	if after := readTree(t, root); !reflect.DeepEqual(before, after) {
		t.Fatal("planner modified workspace inputs")
	}
}

func TestPlanRejectsCrossTypeDuplicateIncludedNames(t *testing.T) {
	root := writeSourceTree(t)
	write(t, filepath.Join(root, "src", "classes", "Main.cls"), "Option Explicit")
	_, err := Plan(Options{Root: root, Config: config.Default()})
	if err == nil || !strings.Contains(err.Error(), "duplicate included VBA component name") {
		t.Fatalf("err = %v", err)
	}
}

func TestPlanAllowsExcludedDuplicateUserForm(t *testing.T) {
	root := writeSourceTree(t)
	write(t, filepath.Join(root, "src", "forms", "alternate", "Login.frm"), "VERSION 5.00\nAttribute VB_Name = \"Login\"\nOption Explicit\n")
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"
	cfg.Build.Exclude = []string{"src/forms/alternate/Login.frm"}

	plan, err := Plan(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if got, want := componentPaths(plan.Included), []string{"src/classes/Service.cls", "src/forms/Login.frm", "src/modules/Main.bas", "src/modules/Tests/TestMain.bas", "src/workbook/ThisWorkbook.bas"}; !reflect.DeepEqual(got, want) {
		t.Fatalf("included = %#v, want %#v", got, want)
	}
	if got, want := componentPaths(plan.Excluded), []string{"src/forms/alternate/Login.frm"}; !reflect.DeepEqual(got, want) {
		t.Fatalf("excluded = %#v, want %#v", got, want)
	}
}

func TestPlanMatchesPersistedSpecsInFRMMode(t *testing.T) {
	root := writeSourceTree(t)
	write(t, filepath.Join(root, "src", "forms", "specs", "Login.yaml"), "kind: xlflow.userform\n")
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"
	cfg.Build.Exclude = []string{"src/forms/specs/Login.yaml"}

	plan, err := Plan(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if got, want := componentPaths(plan.Excluded), []string{"src/forms/Login.frm"}; !reflect.DeepEqual(got, want) {
		t.Fatalf("excluded = %#v, want %#v", got, want)
	}
	if got := plan.Excluded[0].RelatedPaths; !reflect.DeepEqual(got, []string{"src/forms/Login.frx", "src/forms/specs/Login.yaml"}) {
		t.Fatalf("form related paths = %#v", got)
	}
}

func TestPlanRecordsEveryMatchingExcludePattern(t *testing.T) {
	root := writeSourceTree(t)
	cfg := config.Default()
	cfg.Build.Exclude = []string{"src/forms/Login.frm", "src/forms/code/Login.bas"}

	plan, err := Plan(Options{Root: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(plan.Warnings) != 0 {
		t.Fatalf("warnings = %#v", plan.Warnings)
	}
	if got, want := componentPaths(plan.Excluded), []string{"src/forms/Login.frm"}; !reflect.DeepEqual(got, want) {
		t.Fatalf("excluded = %#v, want %#v", got, want)
	}
}

func TestPlanRejectsInvalidLayoutsAndPaths(t *testing.T) {
	t.Run("invalid pattern", func(t *testing.T) {
		cfg := config.Default()
		cfg.Build.Exclude = []string{"../outside/**"}
		_, err := Plan(Options{Root: writeSourceTree(t), Config: cfg})
		if err == nil || !strings.Contains(err.Error(), "project-root-relative") {
			t.Fatalf("err = %v", err)
		}
	})
	t.Run("Windows absolute pattern", func(t *testing.T) {
		cfg := config.Default()
		cfg.Build.Exclude = []string{"C:\\repo\\src\\modules\\Tests\\**"}
		_, err := Plan(Options{Root: writeSourceTree(t), Config: cfg})
		if err == nil || !strings.Contains(err.Error(), "project-root-relative") {
			t.Fatalf("err = %v", err)
		}
	})
	t.Run("same base and output", func(t *testing.T) {
		root := writeSourceTree(t)
		_, err := Plan(Options{Root: root, Config: config.Default(), OutputPath: "build/Book.xlsm"})
		if err == nil || !strings.Contains(err.Error(), "different files") {
			t.Fatalf("err = %v", err)
		}
	})
	t.Run("missing source root", func(t *testing.T) {
		root := writeSourceTree(t)
		cfg := config.Default()
		cfg.Src.Classes = "src/no-classes"
		_, err := Plan(Options{Root: root, Config: cfg})
		if err == nil || !strings.Contains(err.Error(), "read class source root") {
			t.Fatalf("err = %v", err)
		}
	})
	t.Run("orphan sidecar", func(t *testing.T) {
		root := writeSourceTree(t)
		write(t, filepath.Join(root, "src", "forms", "code", "Missing.bas"), "Option Explicit")
		_, err := Plan(Options{Root: root, Config: config.Default()})
		if err == nil || !strings.Contains(err.Error(), "orphan UserForm sidecar") {
			t.Fatalf("err = %v", err)
		}
	})
	t.Run("unsupported source", func(t *testing.T) {
		root := writeSourceTree(t)
		write(t, filepath.Join(root, "src", "modules", "notes.txt"), "not VBA")
		_, err := Plan(Options{Root: root, Config: config.Default()})
		if err == nil || !strings.Contains(err.Error(), "unsupported standard source file") {
			t.Fatalf("err = %v", err)
		}
	})
}

func writeSourceTree(t *testing.T) string {
	t.Helper()
	root := t.TempDir()
	write(t, filepath.Join(root, "src", "modules", "Main.bas"), "Option Explicit")
	write(t, filepath.Join(root, "src", "modules", "Tests", "TestMain.bas"), "Option Explicit")
	write(t, filepath.Join(root, "src", "classes", "Service.cls"), "Option Explicit")
	write(t, filepath.Join(root, "src", "workbook", "ThisWorkbook.bas"), "Option Explicit")
	write(t, filepath.Join(root, "src", "forms", "Login.frm"), "VERSION 5.00\nAttribute VB_Name = \"Login\"\nOption Explicit\n")
	write(t, filepath.Join(root, "src", "forms", "Login.frx"), "binary")
	write(t, filepath.Join(root, "src", "forms", "code", "Login.bas"), "Option Explicit\n")
	return root
}

func write(t *testing.T, path, body string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func componentPaths(components []BuildComponent) []string {
	paths := make([]string, len(components))
	for i, component := range components {
		paths[i] = component.SourcePath
	}
	return paths
}

func readTree(t *testing.T, root string) map[string]string {
	t.Helper()
	out := map[string]string{}
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		rel, _ := filepath.Rel(root, path)
		out[filepath.ToSlash(rel)] = string(body)
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	return out
}
