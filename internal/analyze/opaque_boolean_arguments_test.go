package analyze

import (
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA248DetectsMultiplePositionalBooleanLiterals(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ProcessData(Optional ByVal first As Boolean = False, Optional ByVal second As Boolean = False, Optional ByVal overwrite As Boolean = False)
End Sub

Public Sub Save(Optional ByVal overwrite As Boolean = False)
End Sub

Public Sub Mixed(Optional ByVal first As Boolean = False, Optional ByVal amount As Long = 0, Optional ByVal last As Boolean = False)
End Sub

Public Sub Run()
  ProcessData True, False
  ProcessData True, second:=False
  ProcessData first:=True, second:=False
  ProcessData enabled, enabled
  UnknownOperation True, False
  ProcessData True
  Mixed True, 1, False
  Save True
  ' xlflow:disable-next-line VBA248
  ProcessData True, False
End Sub
`)
	cfg := config.Default()
	if findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run(); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(findings, "VBA248"); len(got) != 0 {
		t.Fatalf("VBA248 is not opt-in; findings = %+v", got)
	}
	cfg.Analyze.DetectOpaqueBooleanArguments = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA248")
	if len(got) != 4 {
		t.Fatalf("VBA248 findings = %+v, want multiple-literal and optional-switch calls", got)
	}
	if got[0].OpaqueBoolean == nil || got[0].OpaqueBoolean.PositionalLiteralCount != 2 || got[0].OpaqueBoolean.NamedArgumentCount != 0 ||
		!reflect.DeepEqual(got[0].OpaqueBoolean.ParameterNames, []string{"first", "second"}) {
		t.Fatalf("unexpected VBA248 context: %+v", got[0])
	}
	if got[0].Procedure != "Run" || got[0].File != "src/modules/Main.bas" {
		t.Fatalf("unexpected VBA248 location: %+v", got[0])
	}
	if got[1].OpaqueBoolean == nil || got[1].OpaqueBoolean.PositionalLiteralCount != 2 {
		t.Fatalf("unexpected second VBA248 context: %+v", got[1])
	}
	if got[2].OpaqueBoolean == nil || got[2].OpaqueBoolean.PositionalLiteralCount != 1 || got[2].OpaqueBoolean.OptionalBooleanParameterCount == nil || *got[2].OpaqueBoolean.OptionalBooleanParameterCount != 3 {
		t.Fatalf("unexpected single-literal VBA248 context: %+v", got[2])
	}
	if got[3].OpaqueBoolean == nil || !reflect.DeepEqual(got[3].OpaqueBoolean.ParameterNames, []string{"first", "last"}) {
		t.Fatalf("unexpected mixed-parameter VBA248 context: %+v", got[3])
	}
}

func TestVBA248ContextPreservesResolvedZeroOptionalCountAndLiteralValues(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Required(first As Boolean, second As Boolean)
End Sub

Public Sub Run()
  Required False, True
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectOpaqueBooleanArguments = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA248")
	if len(got) != 1 || got[0].OpaqueBoolean == nil {
		t.Fatalf("VBA248 findings = %+v, want one finding with context", got)
	}
	context := got[0].OpaqueBoolean
	if context.OptionalBooleanParameterCount == nil || *context.OptionalBooleanParameterCount != 0 {
		t.Fatalf("resolved required Boolean parameters lost zero optional count: %+v", context)
	}
	if !strings.Contains(got[0].Suggestion, "first:=False, second:=True") {
		t.Fatalf("suggestion did not preserve literal values: %q", got[0].Suggestion)
	}
	encoded, err := json.Marshal(got[0])
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"optional_boolean_parameter_count":0`) {
		t.Fatalf("resolved zero optional count missing from JSON: %s", encoded)
	}
}

func TestVBA248NamedArgumentsSuppressSingleLiteralAndBatchRealtimeMatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ProcessData(Optional ByVal first As Boolean = False, Optional ByVal second As Boolean = False)
End Sub

Public Sub Run()
  ProcessData True, second:=False
  ProcessData first:=True, second:=False
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectOpaqueBooleanArguments = true
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(batch, "VBA248"); len(got) != 0 {
		t.Fatalf("named arguments did not suppress single literal: %+v", got)
	}
	if got := findingsByCode(realtime, "VBA248"); len(got) != 0 {
		t.Fatalf("realtime named-argument findings = %+v", got)
	}
	if got, want := findingsByCode(realtime, "VBA248"), findingsByCode(batch, "VBA248"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA248 mismatch: batch=%+v realtime=%+v", want, got)
	}
}
