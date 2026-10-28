package analyze

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func TestConditionalProjectionSkipsInactiveByRefCalls(t *testing.T) {
	t.Parallel()
	// vba-web/WebHelpers.bas: EnableCustomFormatting=True excludes both
	// LogWarning calls in the #Else branches, even though active calls with
	// literal arguments still qualify for the runtime-safety policy.
	source := `Option Explicit
#Const EnableCustomFormatting = True
Private Sub LogWarning(ByRef Message As String)
    Message = Message & "!"
End Sub
Public Sub Run()
#If EnableCustomFormatting Then
    Debug.Print "enabled"
#Else
    LogWarning "inactive"
#End If
    LogWarning "active"
End Sub
`
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	cfg.Analyze.DetectByRefArgumentMismatch = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	wantLine := 1 + strings.Count(source[:strings.Index(source, `    LogWarning "active"`)], "\n")
	got := findingsByCode(findings, "VBA206")
	if len(got) != 1 || got[0].Line != wantLine {
		t.Fatalf("VBA206 findings = %+v, want only the active call at line %d", got, wantLine)
	}
	diagnostics := (intel.Analyzer{RootDir: dir, Config: cfg}).Diagnostics(intel.Document{Path: filepath.Join(dir, "src", "modules", "Main.bas"), Source: source})
	var lines []int
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == "VBA206" {
			lines = append(lines, diagnostic.Range.Start.Line+1)
		}
	}
	if len(lines) != 1 || lines[0] != wantLine {
		t.Fatalf("LSP VBA206 lines = %v, want only %d", lines, wantLine)
	}
}
