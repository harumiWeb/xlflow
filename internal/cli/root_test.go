package cli

import (
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/xuri/excelize/v2"
)

func TestRootCommandIncludesTestCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"test"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "test" {
		t.Fatalf("expected test command, got %#v", cmd)
	}
	if flag := cmd.Flags().Lookup("filter"); flag == nil {
		t.Fatal("expected test command to define --filter")
	}
}

func TestRootCommandIncludesRunFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"arg", "input", "save", "save-as", "trace"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected run command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesTraceInjectCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"trace", "inject"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "inject" {
		t.Fatalf("expected trace inject command, got %#v", cmd)
	}
}

func TestRootCommandIncludesDiffCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"diff"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "diff" {
		t.Fatalf("expected diff command, got %#v", cmd)
	}
	for _, name := range []string{"vba-before", "vba-after"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected diff command to define --%s", name)
		}
	}
}

func TestBuildDiffOptionsRejectsPartialVBADirs(t *testing.T) {
	_, err := buildDiffOptions(".", "before.xlsx", "after.xlsx", "src1", "")
	if err == nil || !strings.Contains(err.Error(), "--vba-before and --vba-after") {
		t.Fatalf("expected vba dir pair error, got %v", err)
	}
}

func TestBuildDiffOptionsRejectsUnsupportedWorkbookExtensions(t *testing.T) {
	_, err := buildDiffOptions(".", "before.xls", "after.xlsx", "", "")
	if err == nil || !strings.Contains(err.Error(), "unsupported extension") {
		t.Fatalf("expected extension error, got %v", err)
	}
}

func TestDiffCommandReturnsSuccessWhenDifferencesExist(t *testing.T) {
	dir := t.TempDir()
	beforePath := filepath.Join(dir, "before.xlsx")
	afterPath := filepath.Join(dir, "after.xlsx")
	before := excelize.NewFile()
	if err := before.SetCellValue("Sheet1", "A1", "old"); err != nil {
		t.Fatal(err)
	}
	if err := before.SaveAs(beforePath); err != nil {
		t.Fatal(err)
	}
	after := excelize.NewFile()
	if err := after.SetCellValue("Sheet1", "A1", "new"); err != nil {
		t.Fatal(err)
	}
	if err := after.SaveAs(afterPath); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"diff", beforePath, afterPath})
	if err := root.Execute(); err != nil {
		t.Fatalf("diff command error = %v, exit = %d", err, output.ExitCode(err))
	}
}

func TestBuildRunOptionsRejectsConflictingSaveFlags(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:hello"}, true, "build\\result.xlsm", false)
	if err == nil || !strings.Contains(err.Error(), "--save and --save-as cannot be combined") {
		t.Fatalf("expected save conflict error, got %v", err)
	}
}

func TestBuildRunOptionsParsesTypedArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "", "fixtures\\Book.xlsm", []string{"string:hello", "int:7", "bool:true"}, false, "", true)
	if err != nil {
		t.Fatal(err)
	}

	want := []excel.RunArgument{
		{Type: "string", Value: "hello"},
		{Type: "int", Value: "7"},
		{Type: "bool", Value: "true"},
	}
	if opts.Macro != "Main.Run" {
		t.Fatalf("macro = %q, want Main.Run", opts.Macro)
	}
	if opts.WorkbookPath != "fixtures\\Book.xlsm" {
		t.Fatalf("workbook path = %q", opts.WorkbookPath)
	}
	if !opts.Trace {
		t.Fatal("expected trace flag to be enabled")
	}
	if !reflect.DeepEqual(opts.Args, want) {
		t.Fatalf("run args = %#v, want %#v", opts.Args, want)
	}
}

func TestBuildRunOptionsAllowsEmptyStringArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:"}, false, "", false)
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.Args) != 1 || opts.Args[0].Type != "string" || opts.Args[0].Value != "" {
		t.Fatalf("run args = %#v", opts.Args)
	}
}

func TestBuildRunOptionsRejectsMalformedTypedArguments(t *testing.T) {
	cfg := config.Default()
	tests := []struct {
		literal string
		wantErr string
	}{
		{"int:not-a-number", "int values must parse"},
		{"bool:maybe", "bool values must be true or false"},
		{"hello", "expected type:value"},
		{"float:3.14", "unsupported --arg type prefix"},
		{"int:", "int values cannot be empty"},
		{"bool:", "bool values cannot be empty"},
	}
	for _, tt := range tests {
		t.Run(tt.literal, func(t *testing.T) {
			_, err := buildRunOptions(cfg, "Main.Run", "", []string{tt.literal}, false, "", false)
			if err == nil {
				t.Fatalf("expected %q to fail", tt.literal)
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %v, want substring %q", err, tt.wantErr)
			}
		})
	}
}
