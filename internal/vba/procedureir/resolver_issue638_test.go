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
	seen := map[string]bool{}
	for _, procedure := range resolved.Procedures {
		for _, access := range procedure.Accesses {
			switch access.Name {
			case "xlCenter", "xlThin", "xlContinuous":
				seen[access.Name] = true
				if access.Resolution.Status == ResolutionAmbiguous {
					t.Fatalf("TypeLib constant %s resolved ambiguously: %#v", access.Name, access.Resolution)
				}
			}
		}
	}
	for _, name := range []string{"xlCenter", "xlThin", "xlContinuous"} {
		if !seen[name] {
			t.Fatalf("source did not retain enum-member access for %s: %#v", name, resolved.Procedures)
		}
	}
	for _, diagnostic := range Diagnostics(resolved, true) {
		if diagnostic.Code == "VB053" {
			t.Fatalf("TypeLib constants produced VB053: %#v", diagnostic)
		}
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
