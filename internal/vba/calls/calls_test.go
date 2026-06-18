package calls

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestInspectExtractsRepresentativeCallSites(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	classDir := filepath.Join(dir, "src", "classes")
	formDir := filepath.Join(dir, "src", "forms")
	for _, path := range []string{moduleDir, classDir, formDir} {
		if err := os.MkdirAll(path, 0o755); err != nil {
			t.Fatal(err)
		}
	}
	main := `Attribute VB_Name = "Main"
Option Explicit
Public Sub RunReport()
    BuildReport 1, 2
    Call SaveReport(Verbose:=True)
    ParenthesizedCall(1, 2)
    result = CalculateTotal(1, 2)
    Debug.Print result
    obj.DoSomething result
    Application.WorksheetFunction.Sum(values)
    Set item = New Customer
    CommandButton1_Click
End Sub

Public Function CalculateTotal(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    CalculateTotal = leftValue + rightValue
End Function
`
	report := `Attribute VB_Name = "ReportBuilder"
Option Explicit
Public Sub BuildReport(ByVal first As Long, ByVal second As Long)
End Sub
Public Sub SaveReport(Optional ByVal Verbose As Boolean = False)
End Sub
Public Sub ParenthesizedCall(ByVal first As Long, ByVal second As Long)
End Sub
`
	customer := `VERSION 1.0 CLASS
Attribute VB_Name = "Customer"
Option Explicit
`
	form := `VERSION 5.00
Begin {00000000-0000-0000-0000-000000000000} UserForm1
End
Attribute VB_Name = "UserForm1"
Option Explicit
Public Sub CommandButton1_Click()
End Sub
`
	mustWrite(t, filepath.Join(moduleDir, "Main.bas"), main)
	mustWrite(t, filepath.Join(moduleDir, "ReportBuilder.bas"), report)
	mustWrite(t, filepath.Join(classDir, "Customer.cls"), customer)
	mustWrite(t, filepath.Join(formDir, "UserForm1.frm"), form)

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 4 {
		t.Fatalf("files = %d, want 4", result.Summary.Files)
	}
	assertCall(t, result.Calls, "BuildReport", "matched", 2)
	save := assertCall(t, result.Calls, "SaveReport", "matched", 1)
	if len(save.Arguments.Named) != 1 || save.Arguments.Named[0].Name != "Verbose" || save.Arguments.Named[0].ValueText != "True" {
		t.Fatalf("unexpected named arguments: %+v", save.Arguments.Named)
	}
	assertCall(t, result.Calls, "ParenthesizedCall", "matched", 2)
	assertCall(t, result.Calls, "CalculateTotal", "matched", 2)
	debug := assertCall(t, result.Calls, "Debug.Print", "external", 1)
	if debug.Callee.Receiver == nil || *debug.Callee.Receiver != "Debug" || debug.Callee.Member != "Print" {
		t.Fatalf("unexpected Debug.Print callee: %+v", debug.Callee)
	}
	assertCall(t, result.Calls, "obj.DoSomething", "member_call", 1)
	assertCall(t, result.Calls, "Application.WorksheetFunction.Sum", "external", 1)
	assertCall(t, result.Calls, "New Customer", "unresolved", 0)
	eventCall := assertCall(t, result.Calls, "CommandButton1_Click", "matched", 0)
	if eventCall.Caller == nil || eventCall.Caller.QualifiedName != "Main.RunReport" {
		t.Fatalf("unexpected caller: %+v", eventCall.Caller)
	}
}

func TestInspectFiltersFromAndTo(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Public Sub First()
    Target
    Other
End Sub
Public Sub Second()
    Target
End Sub
Public Sub Target()
End Sub
Public Sub Other()
End Sub
`
	mustWrite(t, filepath.Join(moduleDir, "Main.bas"), body)

	result, err := Inspect(Options{RootDir: dir, Config: cfg, From: "Main.First", To: "Target"})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Calls) != 1 {
		t.Fatalf("calls = %+v, want one filtered call", result.Calls)
	}
	if result.Calls[0].Caller == nil || result.Calls[0].Caller.Name != "First" || result.Calls[0].Callee.BaseName != "Target" {
		t.Fatalf("unexpected filtered call: %+v", result.Calls[0])
	}
}

func TestInspectReportsParseRecoveryWithoutCrashing(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Broken"
Option Explicit
Public Function Broken(ByVal value As String
    Foo
End Function
`
	mustWrite(t, filepath.Join(moduleDir, "Broken.bas"), body)

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 1 || result.Summary.ParseErrors != 1 || result.Summary.MissingNodes != 1 {
		t.Fatalf("unexpected recovery summary: %+v", result.Summary)
	}
}

func mustWrite(t *testing.T, path, body string) {
	t.Helper()
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func assertCall(t *testing.T, calls []Call, text, status string, argCount int) Call {
	t.Helper()
	for _, call := range calls {
		if call.Callee.Text == text {
			if call.Resolution.Status != status || call.Arguments.Count != argCount {
				t.Fatalf("call %s = status %s args %d, want %s/%d: %+v", text, call.Resolution.Status, call.Arguments.Count, status, argCount, call)
			}
			if call.Range.StartLine == 0 || call.File == "" || call.Module == "" {
				t.Fatalf("call %s missing location context: %+v", text, call)
			}
			return call
		}
	}
	t.Fatalf("missing call %q in %+v", text, calls)
	return Call{}
}
