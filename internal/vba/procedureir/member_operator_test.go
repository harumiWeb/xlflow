package procedureir

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestQualifiedMemberOperatorPreservesDotAndBang(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Members.bas"}, []byte(`Public Sub Run()
  Dim value!
  Dim rs As Object
  value = 1
  rs!CustomerName = "x"
  rs.CustomerName = "y"
  Call rs!Save()
  Call rs.Save()
  Call rs!CustomerName.ToString()
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if doc.Parse.HasError || doc.Parse.HasMissing {
		t.Fatalf("unexpected parse recovery: %+v", doc.Parse)
	}
	procedure := doc.Procedures[0]
	operators := map[string]MemberOperator{}
	for _, expression := range procedure.Expressions {
		if expression.Kind == ExpressionMember {
			operators[expression.Text] = expression.MemberOperator
		}
		if expression.SyntaxKind == "bang_identifier" && expression.MemberOperator != "" {
			t.Fatalf("identifier type character was treated as member operator: %+v", expression)
		}
	}
	if operators["rs!CustomerName"] != MemberOperatorBang || operators["rs.CustomerName"] != MemberOperatorDot {
		t.Fatalf("member expression operators = %#v", operators)
	}
	if len(procedure.Calls) != 3 {
		t.Fatalf("calls = %#v", procedure.Calls)
	}
	if procedure.Calls[0].MemberOperator != MemberOperatorBang || procedure.Calls[1].MemberOperator != MemberOperatorDot {
		t.Fatalf("call operators = %#v", procedure.Calls)
	}
	if procedure.Calls[2].MemberOperator != MemberOperatorDot {
		t.Fatalf("nested call operator = %#v", procedure.Calls[2])
	}
	if operators["rs!CustomerName.ToString"] != MemberOperatorDot || operators["rs!CustomerName"] != MemberOperatorBang {
		t.Fatalf("nested member expression operators = %#v", operators)
	}

	encoded, err := json.Marshal(doc)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(encoded), "memberOperator") {
		t.Fatalf("internal member operator leaked into wire JSON: %s", encoded)
	}
}

func TestQualifiedMemberOperatorSurvivesCloneRebaseAndMaterialize(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Members.bas"}, []byte(`Public Sub Run()
  Dim rs As Object
  rs!CustomerName = "x"
  Call rs!Save()
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	want := doc.Procedures[0]
	assertMemberOperator := func(label string, procedure ProcedureIR) {
		var expression MemberOperator
		for _, candidate := range procedure.Expressions {
			if candidate.Text == "rs!CustomerName" {
				expression = candidate.MemberOperator
				break
			}
		}
		if expression != MemberOperatorBang || len(procedure.Calls) != 1 || procedure.Calls[0].MemberOperator != MemberOperatorBang {
			t.Fatalf("%s lost member operator: expression=%q calls=%#v", label, expression, procedure.Calls)
		}
	}
	assertMemberOperator("source", want)
	assertMemberOperator("clone", CloneProcedureIR(want))

	oldBase := want.Symbol.BodyRange
	newBase := oldBase
	newBase.StartByte += 100
	newBase.EndByte += 100
	newBase.StartLine++
	newBase.EndLine++
	assertMemberOperator("rebase", RebaseProcedure(want, oldBase, newBase))

	view := ResolveView(doc, NewResolver(nil))
	assertMemberOperator("materialize", view.Materialize().Procedures[0])

}
