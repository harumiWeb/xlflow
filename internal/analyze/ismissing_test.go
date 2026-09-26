package analyze

import (
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestInvalidIsMissingUsageFindsMisuseAndAcceptsOptionalVariants(t *testing.T) {
	cfg := config.Default()
	cfg.Analyze.DetectInvalidIsMissingUsage = true
	source := []byte(`Option Explicit
Private mValue As Variant

Private Sub Valid(Optional explicitValue As Variant, Optional implicitValue)
    If IsMissing(explicitValue) Then
    End If
    If IsMissing(((implicitValue))) Then
    End If
    If VBA.IsMissing(implicitValue) Then
    End If
End Sub

Private Sub NonVariant(Optional value As Long)
    If IsMissing(value) Then
    End If
End Sub

Private Sub NonOptional(value As Variant)
    If IsMissing(value) Then
    End If
End Sub

Private Sub DefaultedVariant(Optional value As Variant = 0)
    If IsMissing(value) Then
    End If
End Sub

Private Sub DefaultedNonVariant(Optional value As Long = 0)
    If IsMissing(value) Then
    End If
End Sub

Private Sub LocalVariable(Optional value As Variant)
    Dim localValue As Variant
    If IsMissing(localValue) Then
    End If
End Sub

Private Sub MemberArgument(Optional value As Variant)
    If IsMissing(Me.mValue) Then
    End If
End Sub

Private Sub ExpressionArgument(Optional value As Variant)
    If IsMissing(value + 1) Then
    End If
End Sub

Private Sub UnrelatedArgument(Optional allowed As Variant, required As Variant)
    If IsMissing(required) Then
    End If
End Sub
`)
	dir := t.TempDir()
	writeModule(t, dir, "Main.cls", string(source))
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA283")
	want := map[string]string{
		"NonVariant":          "value",
		"NonOptional":         "value",
		"DefaultedNonVariant": "value",
		"LocalVariable":       "localValue",
		"MemberArgument":      "Me.mValue",
		"ExpressionArgument":  "value + 1",
		"UnrelatedArgument":   "required",
	}
	if len(got) != len(want) {
		t.Fatalf("VBA283 findings = %+v, want %d", got, len(want))
	}
	seen := make(map[string]bool, len(got))
	for _, finding := range got {
		argument, ok := want[finding.Procedure]
		if !ok || seen[finding.Procedure] {
			t.Fatalf("unexpected or duplicate VBA283 finding: %+v", finding)
		}
		seen[finding.Procedure] = true
		if finding.Severity != "warning" {
			t.Errorf("%s severity = %q, want warning", finding.Procedure, finding.Severity)
		}
		line := strings.Split(string(source), "\n")[finding.Line-1]
		if gotColumn := strings.Index(line, argument) + 1; finding.Column != gotColumn {
			t.Errorf("%s column = %d, want argument start %d (%q)", finding.Procedure, finding.Column, gotColumn, argument)
		}
	}
}

func TestInvalidIsMissingUsageRequiresIntrinsicResolution(t *testing.T) {
	cfg := config.Default()
	cfg.Analyze.DetectInvalidIsMissingUsage = true
	findings := runParameterPassingAnalysis(t, cfg, map[string]string{"Main.bas": `Option Explicit
Public Function IsMissing(ByVal value As Variant) As Boolean
    IsMissing = False
End Function

Private Sub Shadowed(Optional value As Long)
    If IsMissing(value) Then
    End If
End Sub

Private Sub ExplicitIntrinsic(Optional value As Long)
    If VBA.IsMissing(value) Then
    End If
End Sub

Private Sub OtherReceiver(Optional value As Long)
    Dim obj As Object
    If obj.IsMissing(value) Then
    End If
End Sub

Private Sub NonCallable(Optional value As Long)
    Dim IsMissing As Variant
    If IsMissing(value) Then
    End If
End Sub
`})
	got := findingsByCode(findings, "VBA283")
	if len(got) != 1 || got[0].Procedure != "ExplicitIntrinsic" {
		t.Fatalf("VBA283 shadowing findings = %+v, want only explicit VBA intrinsic", got)
	}
}

func TestInvalidIsMissingUsageIgnoresIndexedArraysNamedIntrinsic(t *testing.T) {
	cfg := config.Default()
	cfg.Analyze.DetectInvalidIsMissingUsage = true
	findings := runParameterPassingAnalysis(t, cfg, map[string]string{
		"ModuleArray.bas": `Option Explicit
Private IsMissing(0 To 1) As Boolean

Private Sub ModuleArray()
    If IsMissing(0) Then
    End If
End Sub

Private Sub ExplicitIntrinsicBesideModuleArray()
    If VBA.IsMissing(0) Then
    End If
End Sub
`,
		"LocalArray.bas": `Option Explicit

Private Sub LocalArray()
    Dim IsMissing(0 To 1) As Boolean
    If IsMissing(0) Then
    End If
End Sub
`,
	})
	got := findingsByCode(findings, "VBA283")
	if len(got) != 1 || got[0].Procedure != "ExplicitIntrinsicBesideModuleArray" {
		t.Fatalf("VBA283 array shadowing findings = %+v, want only explicit VBA intrinsic", got)
	}
}

func TestInvalidIsMissingUsageIsOptInAndFailsOpenOnWrongArity(t *testing.T) {
	source := []byte(`Option Explicit
Private Sub Run(Optional value As Variant)
    If IsMissing(value, value) Then
    End If
    If IsMissing() Then
    End If
End Sub
`)
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", string(source))
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA283"); len(got) != 0 {
		t.Fatalf("default or malformed IsMissing findings = %+v, want none", got)
	}
	cfg := config.Default()
	cfg.Analyze.DetectInvalidIsMissingUsage = true
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA283"); len(got) != 0 {
		t.Fatalf("wrong-arity IsMissing findings = %+v, want none", got)
	}
}

func TestInvalidIsMissingUsageBatchRealtimeParityAndSuppression(t *testing.T) {
	dir := t.TempDir()
	source := []byte(`Option Explicit
Private Sub Run(Optional value As Long) ' xlflow:disable-next-line VBA283
    If IsMissing(value) Then
    End If
    If IsMissing(value) Then
    End If
End Sub
`)
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	writeModule(t, dir, "Main.bas", string(source))
	cfg := config.Default()
	cfg.Analyze.DetectInvalidIsMissingUsage = true
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	batchFindings := findingsByCode(batch, "VBA283")
	realtimeFindings := findingsByCode(realtime, "VBA283")
	if len(batchFindings) != 1 || len(realtimeFindings) != 1 {
		t.Fatalf("VBA283 batch/realtime findings = %+v / %+v, want one unsuppressed finding each", batchFindings, realtimeFindings)
	}
	if !reflect.DeepEqual(batchFindings, realtimeFindings) {
		t.Fatalf("VBA283 batch/realtime differ:\nbatch: %+v\nrealtime: %+v", batchFindings, realtimeFindings)
	}
}
