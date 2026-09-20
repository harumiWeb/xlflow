package procedureir

import (
	"sort"
	"testing"
)

func TestIsAssignmentTargetCallKeepsSameNameRHSInvocation(t *testing.T) {
	document, err := BuildSource(BuildOptions{Path: "Main.bas", ModuleName: "Main", ModuleKind: "standard"}, []byte(`Option Explicit
Public Function Split(ByVal expression As String) As String()
  Split(0) = Split(1)
End Function
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(document.Procedures) != 1 {
		t.Fatalf("procedures = %d, want one", len(document.Procedures))
	}
	procedure := document.Procedures[0]
	var calls []CallSite
	for _, call := range procedure.Calls {
		if call.Callee.BaseName == "Split" {
			calls = append(calls, call)
		}
	}
	if len(calls) != 2 {
		t.Fatalf("Split calls = %#v, want indexed target and RHS call", procedure.Calls)
	}
	sort.Slice(calls, func(i, j int) bool { return calls[i].Range.StartByte < calls[j].Range.StartByte })
	if !IsAssignmentTargetCall(calls[0], procedure) {
		t.Fatalf("left-hand indexed call was not recognized as assignment target: call=%#v statements=%#v expressions=%#v", calls[0], procedure.Statements, procedure.Expressions)
	}
	if IsAssignmentTargetCall(calls[1], procedure) {
		t.Fatalf("same-name RHS call was incorrectly recognized as assignment target: %#v", calls[1])
	}
}

func TestIsAssignmentTargetCallRecognizesRecoveredCallStatementTarget(t *testing.T) {
	document, err := BuildSource(BuildOptions{Path: "Main.cls", ModuleName: "Main", ModuleKind: "class"}, []byte(`Option Explicit
Public Property Get LocalPaths() As String()
  ReDim LocalPaths(0 To 0)
  LocalPaths(0) = This.Path
End Property
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(document.Procedures) != 1 || len(document.Procedures[0].Calls) != 1 {
		t.Fatalf("procedures/calls = %#v, want one recovered return-slot call", document.Procedures)
	}
	if !IsAssignmentTargetCall(document.Procedures[0].Calls[0], document.Procedures[0]) {
		t.Fatalf("recovered call-statement return slot was not recognized: %#v", document.Procedures[0].Calls[0])
	}
}
