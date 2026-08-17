package procedureir

import "testing"

func TestIssue638TypeLibEnumDuplicatesAreNotAmbiguous(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`Public Sub Consume(ByVal value As Long)
End Sub
Public Sub TestAlignment()
    Dim value As Long
    value = xlCenter
    value = xlThin
    value = xlContinuous
    Consume xlCenter
    Call Consume(xlThin)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolved := Resolve(doc, NewResolver([]ResolverSymbol{
		{Name: "xlCenter", Parent: "XlHAlign", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
		{Name: "xlCenter", Parent: "XlHorizontalAlignment", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
		{Name: "xlThin", Parent: "XlBorderWeight", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
		{Name: "xlThin", Parent: "XlPattern", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
		{Name: "xlContinuous", Parent: "XlLineStyle", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
		{Name: "xlContinuous", Parent: "XlBorderStyle", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
	}))
	seen := map[string]int{}
	for _, procedure := range resolved.Procedures {
		for _, access := range procedure.Accesses {
			switch access.Name {
			case "xlCenter", "xlThin", "xlContinuous":
				seen[access.Name]++
				if access.Resolution.Status != ResolutionExternal {
					t.Fatalf("TypeLib constant %s status = %s, want external: %#v", access.Name, access.Resolution.Status, access.Resolution)
				}
			}
		}
	}
	for name, want := range map[string]int{"xlCenter": 2, "xlThin": 2, "xlContinuous": 1} {
		if got := seen[name]; got != want {
			t.Fatalf("enum-member accesses for %s = %d, want %d: %#v", name, got, want, resolved.Procedures)
		}
	}
	if diagnostics := Diagnostics(resolved, true); len(diagnostics) != 0 {
		t.Fatalf("TypeLib constants produced diagnostics: %#v", diagnostics)
	}
}

func TestIssue638ProjectEnumAmbiguityRemainsProvableWithTypeLibMetadata(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`Public Enum FirstState
    Ready = 1
End Enum
Public Enum SecondState
    Ready = 2
End Enum
Public Sub Run()
    Dim value As Long
    value = Ready
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolved := Resolve(doc, NewResolver([]ResolverSymbol{
		{Name: "Ready", Parent: "FirstState", Module: "Main", ModuleKind: "standard", Kind: "enum_member"},
		{Name: "Ready", Parent: "SecondState", Module: "Main", ModuleKind: "standard", Kind: "enum_member"},
		{Name: "Ready", Parent: "SomeExternalState", Module: "Excel", ModuleKind: "external", Kind: "enum_member"},
	}))
	diagnostics := Diagnostics(resolved, true)
	if len(diagnostics) != 1 || diagnostics[0].Code != "VB053" {
		t.Fatalf("diagnostics = %#v, want one VB053", diagnostics)
	}
}
