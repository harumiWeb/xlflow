package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestCompileEquivalentFindingsIncludesCFGControlFlowLegalityAndPreflight(t *testing.T) {
	source := []byte(`Sub Main()
    GoTo Missing
    Label1:
    label1:
    For i = 1 To 2
    Next wrong
    Exit Function
End Sub
`)
	ir, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Main.bas"}, source)
	if err != nil {
		t.Fatal(err)
	}
	file := parsedFile{
		Path: "Main.bas", Source: source, Lines: normalizedSourceLines(string(source)),
		IR: ir, CFG: vbacfg.BuildDocument(ir),
	}
	findings, preflight := (Analyzer{RootDir: ".", Config: config.Default()}).compileEquivalentFindings(file)
	seen := map[string]int{}
	for _, finding := range findings {
		if finding.Code >= "VB055" && finding.Code <= "VB058" {
			seen[finding.Code]++
		}
	}
	for _, code := range []string{"VB055", "VB056", "VB057", "VB058"} {
		if seen[code] != 1 {
			t.Fatalf("%s count = %d, findings = %+v", code, seen[code], findings)
		}
	}
	for _, finding := range preflight {
		if finding.Code >= "VB055" && finding.Code <= "VB058" && finding.Severity != "error" {
			t.Fatalf("%s preflight severity = %q", finding.Code, finding.Severity)
		}
	}
}
