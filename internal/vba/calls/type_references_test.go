package calls

import "testing"

func TestExtractParsedCollectsTypeReferencesWithoutChangingCallSites(t *testing.T) {
	result, err := ExtractSource(SourceOptions{Path: "Service.cls", ModuleKind: "class"}, []byte(`
Implements IFoo

Public Sub Run()
    Dim service As OrderService
    Dim cached As New CacheService
    Set service = New OrderService
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]int{}
	for _, ref := range result.TypeReferences {
		got[ref.Kind+":"+ref.Target]++
		if ref.File != "Service.cls" || ref.Module != "Service" {
			t.Fatalf("type reference context = %+v", ref)
		}
	}
	for key, want := range map[string]int{
		"implements:IFoo":         1,
		"uses_type:OrderService":  1,
		"uses_type:CacheService":  1,
		"constructs:CacheService": 1,
		"constructs:OrderService": 1,
	} {
		if got[key] != want {
			t.Fatalf("type references = %+v, missing %s=%d", got, key, want)
		}
	}
	if len(result.CallSites) == 0 {
		t.Fatal("New expression compatibility call site was removed")
	}
}
