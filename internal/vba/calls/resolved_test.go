package calls

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestFromResolvedIRPreservesCanonicalStatusesAndCopiesCandidates(t *testing.T) {
	ir := procedureir.DocumentIR{Procedures: []procedureir.ProcedureIR{{Calls: []procedureir.CallSite{
		{Range: ast.Range{StartByte: 20}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionNonCallable}},
		{Range: ast.Range{StartByte: 10}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "Main.Work", Kind: "sub", File: "Main.bas", Line: 5}}}},
		{Range: ast.Range{StartByte: 30}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionIncomplete}},
		{Range: ast.Range{StartByte: 40}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionDynamic}},
	}}}}
	result := FromResolvedIR(ir)
	for index, want := range []string{"matched", "non_callable", "incomplete", "dynamic"} {
		if result[index].Resolution.Status != want {
			t.Errorf("call %d status=%s want%s", index, result[index].Resolution.Status, want)
		}
	}
	result[0].Resolution.Candidates[0].File = "changed.bas"
	if ir.Procedures[0].Calls[1].Resolution.Candidates[0].File != "Main.bas" {
		t.Fatal("projection mutated resolved IR")
	}
}
