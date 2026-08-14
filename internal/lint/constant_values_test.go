package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestConstantValuesFromSourceResolvesConstAndEnumReferences(t *testing.T) {
	source := `Public Const Limit As Long = 2
Public Const Expanded As Long = Limit + 3
Enum State
    Ready = Expanded
    Done
End Enum
Public Const RuntimeValue As Long = Missing()
#If WINDOWS Then
Public Const ConditionalValue As Long = 99
#End If
`
	values := ConstantValuesFromSource(source, nil, nil)
	for name, want := range map[string]int64{
		"limit":       2,
		"expanded":    5,
		"ready":       5,
		"state.ready": 5,
		"done":        6,
		"state.done":  6,
	} {
		value, ok := values[name]
		if !ok || value.Kind != constexpr.ValueLong || value.Integer != want {
			t.Fatalf("value[%q] = %#v, %v; want Long(%d)", name, value, ok, want)
		}
	}
	if _, ok := values["runtimevalue"]; ok {
		t.Fatal("runtime-dependent Const must not be projected as a known value")
	}
	if _, ok := values["conditionalvalue"]; ok {
		t.Fatal("conditional Const must remain unknown")
	}
}

func TestConstantValuesFromSourceNilIRExcludesProcedureLocals(t *testing.T) {
	source := `Public Const ModuleValue As Long = 1
Sub Example()
    Const LocalValue As Long = 2
End Sub
`
	values := ConstantValuesFromSource(source, nil, nil)
	if _, ok := values["modulevalue"]; !ok {
		t.Fatal("module-level Const should remain available without IR")
	}
	if _, ok := values["localvalue"]; ok {
		t.Fatal("procedure-local Const must not be projected without IR")
	}
}

func TestProjectConstantValuesResolvesForwardReferencesAndAmbiguity(t *testing.T) {
	first := procedureir.DocumentIR{
		ModuleName: "First", ModuleKind: "standard",
		Declarations: []procedureir.Declaration{{Name: "FirstValue", Kind: "const", IsConst: true, Visibility: "Public"}},
	}
	second := procedureir.DocumentIR{
		ModuleName: "Second", ModuleKind: "standard",
		Declarations: []procedureir.Declaration{{Name: "SecondValue", Kind: "const", IsConst: true, Visibility: "Public"}},
	}
	values := ProjectConstantValues([]ConstantValueDocument{
		{Source: "Public Const FirstValue As Long = SecondValue\n", IR: &first},
		{Source: "Public Const SecondValue As Long = 3\n", IR: &second},
	}, nil)
	if got, ok := values["first.firstvalue"]; !ok || got.Integer != 3 {
		t.Fatalf("qualified First.FirstValue = %#v, %v; want 3", got, ok)
	}

	duplicate := procedureir.DocumentIR{
		ModuleName: "Duplicate", ModuleKind: "standard",
		Declarations: []procedureir.Declaration{{Name: "SecondValue", Kind: "const", IsConst: true, Visibility: "Public"}},
	}
	values = ProjectConstantValues([]ConstantValueDocument{
		{Source: "Public Const FirstValue As Long = SecondValue\n", IR: &first},
		{Source: "Public Const SecondValue As Long = 3\n", IR: &second},
		{Source: "Public Const SecondValue As Long = 4\n", IR: &duplicate},
	}, nil)
	if _, ok := values["secondvalue"]; ok {
		t.Fatal("duplicate public names must not be projected unqualified")
	}
	if got, ok := values["second.secondvalue"]; !ok || got.Integer != 3 {
		t.Fatalf("qualified Second.SecondValue = %#v, %v; want 3", got, ok)
	}
	if _, ok := values["first.firstvalue"]; ok {
		t.Fatal("FirstValue should remain unresolved while its reference is ambiguous")
	}
}
