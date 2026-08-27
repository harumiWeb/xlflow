package procedureir

import "testing"

type genericProcedureResolver struct{}

func (genericProcedureResolver) ResolveCall(CallSite) CallResolution {
	return CallResolution{Status: ResolutionMatched}
}

func (genericProcedureResolver) ResolveSymbol(SymbolReference) SymbolResolution {
	return SymbolResolution{Scope: ScopeProject, Status: ResolutionMatched}
}

type countingEventResolver struct {
	SymbolResolver
	eventCalls int
}

func (r *countingEventResolver) ResolveEvent(ref SymbolReference) SymbolResolution {
	r.eventCalls++
	return r.SymbolResolver.ResolveEvent(ref)
}

func TestIssue590ResolverNegativeOutcomes(t *testing.T) {
	r := NewResolver([]ResolverSymbol{
		{Name: "value", Module: "Main", ModuleKind: "standard", Kind: "const", Visibility: "Public", File: "Main.bas", Line: 1},
		{Name: "Run", Module: "Main", ModuleKind: "standard", Kind: "sub", Visibility: "Public", File: "Main.bas", Line: 2},
	})
	got := r.ResolveCall(CallSite{
		Caller: ProcedureRef{Name: "Caller", QualifiedName: "Main.Caller"},
		Callee: Callee{Text: "value", BaseName: "value", Member: "value"},
	})
	if got.Status != ResolutionNonCallable || !got.ProjectLocal {
		t.Fatalf("non-callable resolution = %#v", got)
	}
	receiver := "Main"
	got = r.ResolveCall(CallSite{
		Caller: ProcedureRef{Name: "Caller", QualifiedName: "Main.Caller"},
		Callee: Callee{Text: "Main.Missing", BaseName: "Missing", Member: "Missing", Receiver: &receiver},
	})
	if got.Status != ResolutionUnresolved || !got.ProjectLocal {
		t.Fatalf("qualified missing resolution = %#v", got)
	}
	got = r.ResolveCall(CallSite{
		Caller: ProcedureRef{Name: "Caller", QualifiedName: "Main.Caller"},
		Callee: Callee{Text: "Missing", BaseName: "Missing", Member: "Missing"},
	})
	if got.Status != ResolutionUnresolved || got.ProjectLocal {
		t.Fatalf("bare missing resolution = %#v", got)
	}
	application := "Application"
	if got := r.ResolveCall(CallSite{Callee: Callee{Text: "Application.Run", BaseName: "Run", Member: "Run", Receiver: &application}}); got.Status != ResolutionDynamic {
		t.Fatalf("dynamic resolution = %#v", got)
	}
	projectReceiver := "customer"
	classResolver := NewResolver([]ResolverSymbol{
		{Name: "customer", Type: "Widget", Module: "Main", ModuleKind: "standard", Kind: "variable", Visibility: "Private"},
		{Name: "Run", Module: "Widget", ModuleKind: "class", Kind: "sub", Visibility: "Public"},
	})
	if got := classResolver.ResolveCall(CallSite{Module: "Main", Callee: Callee{Text: "customer.Missing", BaseName: "Missing", Member: "Missing", Receiver: &projectReceiver}}); got.Status != ResolutionUnresolved || !got.ProjectLocal {
		t.Fatalf("typed project receiver resolution = %#v", got)
	}
	if got := classResolver.ResolveCall(CallSite{Module: "Main", Callee: Callee{Text: "customer.Run", BaseName: "Run", Member: "Run", Receiver: &projectReceiver}}); got.Status != ResolutionMatched || len(got.Candidates) != 1 {
		t.Fatalf("typed project receiver match = %#v", got)
	}
	formResolver := NewResolver([]ResolverSymbol{{Name: "Picker", Module: "Picker", ModuleKind: "form", Kind: "sub"}})
	me := "Me"
	if got := formResolver.ResolveCall(CallSite{Module: "Picker", Callee: Callee{Text: "Me.Show", BaseName: "Show", Member: "Show", Receiver: &me}}); got.Status != ResolutionExternal {
		t.Fatalf("form host member resolution = %#v", got)
	}
	private := NewResolver([]ResolverSymbol{{
		Name: "Hidden", Module: "Other", ModuleKind: "standard", Kind: "sub", Visibility: "Private",
	}})
	if got := private.ResolveCall(CallSite{Module: "Main", Caller: ProcedureRef{QualifiedName: "Main.Run"}, Callee: Callee{Text: "Hidden", BaseName: "Hidden", Member: "Hidden"}}); got.Status != ResolutionNonCallable || !got.ProjectLocal {
		t.Fatalf("cross-module private resolution = %#v", got)
	}
}

func TestProcedureOnlyResolverSharesCanonicalCandidateStorage(t *testing.T) {
	full := NewResolver([]ResolverSymbol{
		{Name: "Helper", Module: "Main", ModuleKind: "standard", Kind: "sub", Visibility: "Public"},
		{Name: "Value", Module: "Main", ModuleKind: "standard", Kind: "const", Visibility: "Public"},
	})
	procedure, ok := ProcedureOnlyResolver(full).(SymbolResolver)
	if !ok {
		t.Fatal("ProcedureOnlyResolver did not return the canonical resolver view")
	}
	if !procedure.procedureOnly {
		t.Fatal("procedure-only resolver view flag was not set")
	}
	fullEntries := full.byName["helper"]
	procedureEntries := procedure.byName["helper"]
	if len(fullEntries) != 1 || len(procedureEntries) != 1 || &fullEntries[0] != &procedureEntries[0] {
		t.Fatal("procedure-only resolver copied canonical candidate storage")
	}
	call := CallSite{Module: "Main", Caller: ProcedureRef{QualifiedName: "Main.Run"}, Callee: Callee{Text: "Value", BaseName: "Value"}}
	if got := full.ResolveCall(call); got.Status != ResolutionNonCallable {
		t.Fatalf("full resolver const call = %#v, want non-callable", got)
	}
	if got := procedure.ResolveCall(call); got.Status != ResolutionUnresolved {
		t.Fatalf("procedure-only const call = %#v, want unresolved", got)
	}
}

func TestIssue590LocalDeclarationShadowsProcedure(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`Public Sub Run()
    Dim Helper As Long
    Helper
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolved := Resolve(doc, NewResolver([]ResolverSymbol{{
		Name: "Helper", Module: "Other", ModuleKind: "standard", Kind: "sub", Visibility: "Public",
	}}))
	if len(resolved.Procedures) != 1 || len(resolved.Procedures[0].Calls) != 1 {
		t.Fatalf("calls = %#v", resolved.Procedures)
	}
	call := resolved.Procedures[0].Calls[0]
	if call.Resolution.Status != ResolutionNonCallable || !call.Resolution.ProjectLocal {
		t.Fatalf("shadowed call = %#v", call.Resolution)
	}
}

func TestIssue590ExternalLibraryQualifierIsNotProjectLocalProof(t *testing.T) {
	r := NewResolverWithCompleteness([]ResolverSymbol{{
		Name: "vbCrLf", Module: "VBA", ModuleKind: "external", Kind: "enum_member", Recovered: true,
	}}, true)
	receiver := "VBA"
	got := r.ResolveCall(CallSite{
		Module: "Main",
		Callee: Callee{Text: "VBA.IsObject", BaseName: "IsObject", Member: "IsObject", Receiver: &receiver},
	})
	if got.Status != ResolutionMemberCall || got.ProjectLocal {
		t.Fatalf("external library call resolution = %#v", got)
	}
}

func TestIssue590EnumResolutionUsesLexicalWinner(t *testing.T) {
	r := NewResolver([]ResolverSymbol{
		{Name: "Ready", Parent: "LocalState", Module: "Main", ModuleKind: "standard", Kind: "enum_member", File: "Main.bas", Line: 2},
		{Name: "Ready", Parent: "ExternalState", Module: "Other", ModuleKind: "standard", Kind: "enum_member", File: "Other.bas", Line: 2},
	})
	local := r.ResolveEnumMember(EnumMemberReference{Name: "Ready", Caller: ProcedureRef{QualifiedName: "Main.Run"}})
	if local.Status != ResolutionMatched || len(local.Candidates) != 1 || local.Candidates[0].QualifiedName != "Main.Ready" {
		t.Fatalf("lexical enum resolution = %#v", local)
	}
	qualified := r.ResolveEnumMember(EnumMemberReference{Name: "Ready", Enum: "ExternalState", Caller: ProcedureRef{QualifiedName: "Main.Run"}})
	if qualified.Status != ResolutionMatched || len(qualified.Candidates) != 1 || qualified.Candidates[0].QualifiedName != "Other.Ready" {
		t.Fatalf("qualified enum resolution = %#v", qualified)
	}
	access := r.ResolveSymbol(SymbolReference{Name: "Ready", Caller: ProcedureRef{QualifiedName: "Main.Run"}})
	if access.Status != ResolutionMatched || len(access.Candidates) != 1 || access.Candidates[0].QualifiedName != "Main.Ready" {
		t.Fatalf("symbol enum resolution = %#v", access)
	}
}

func TestIssue590DiagnosticsReportAmbiguousModuleEnumMember(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`
Public Enum FirstState
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
		{Name: "Ready", Parent: "FirstState", Module: "Main", ModuleKind: "standard", Kind: "enum_member", File: "Main.bas", Line: 2},
		{Name: "Ready", Parent: "SecondState", Module: "Main", ModuleKind: "standard", Kind: "enum_member", File: "Main.bas", Line: 5},
	}))
	diagnostics := Diagnostics(resolved, true)
	if len(diagnostics) != 1 || diagnostics[0].Code != "VB053" {
		t.Fatalf("diagnostics = %#v, want one VB053", diagnostics)
	}
}

func TestIssue590IRRetainsEnumMembersAndRaiseEvents(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Widget.cls", ModuleKind: "class"}, []byte(`Public Enum State
    Ready = 1
    Busy
End Enum
Public Event Changed(ByVal value As Long)
Public Sub Run()
    RaiseEvent Changed(1)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	var members []Declaration
	for _, declaration := range doc.Declarations {
		if declaration.Kind == "enum_member" {
			members = append(members, declaration)
		}
	}
	if len(members) != 2 || members[0].Parent != "State" || members[0].Range.StartByte >= members[0].Range.EndByte {
		t.Fatalf("enum members = %#v", members)
	}
	if len(doc.Procedures) != 1 || len(doc.Procedures[0].RaiseEvents) != 1 {
		t.Fatalf("raise events = %#v", doc.Procedures)
	}
	ref := doc.Procedures[0].RaiseEvents[0]
	if ref.Name != "Changed" || ref.Range.StartByte >= ref.Range.EndByte || ref.Arguments.Count != 1 {
		t.Fatalf("raise event reference = %#v", ref)
	}
}

func TestIssue590DiagnosticsReportUndeclaredRaiseEvent(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Widget.cls", ModuleKind: "class"}, []byte(`Public Sub Run()
    RaiseEvent Missing(1)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolved := Resolve(doc, NewResolver([]ResolverSymbol{
		{Name: "Changed", Module: "Widget", ModuleKind: "class", Kind: "event", Visibility: "Public"},
	}))
	diagnostics := Diagnostics(resolved, true)
	if len(diagnostics) != 1 || diagnostics[0].Code != "VB054" {
		t.Fatalf("diagnostics = %#v, want one VB054", diagnostics)
	}
	if diagnostics[0].Range.StartByte >= diagnostics[0].Range.EndByte {
		t.Fatalf("VB054 range = %#v, want event identifier range", diagnostics[0].Range)
	}
}

func TestIssue590RaiseEventFallbackIsIncomplete(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Widget.cls", ModuleKind: "class"}, []byte(`Public Event Changed()
Public Sub Run()
    RaiseEvent Changed
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolved := Resolve(doc, genericProcedureResolver{})
	if got := resolved.Procedures[0].RaiseEvents[0].Resolution.Status; got != ResolutionIncomplete {
		t.Fatalf("RaiseEvent fallback status = %q, want incomplete", got)
	}
	if got := resolved.Procedures[0].Calls[0].Resolution.Status; got != ResolutionIncomplete {
		t.Fatalf("RaiseEvent call fallback status = %q, want incomplete", got)
	}
	if diagnostics := Diagnostics(resolved, true); len(diagnostics) != 0 {
		t.Fatalf("incomplete RaiseEvent diagnostics = %#v, want none", diagnostics)
	}
}

func TestIssue590RaiseEventArgumentCountRetained(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Widget.cls", ModuleKind: "class"}, []byte(`Public Event Changed(ByVal first As Long, ByVal second As Long)
Public Sub Run()
    RaiseEvent Changed(1, 2)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if got := doc.Procedures[0].RaiseEvents[0].Arguments.Count; got != 2 {
		t.Fatalf("RaiseEvent argument count = %d, want 2", got)
	}
}

func TestIssue590RaiseEventResolutionIsNotDuplicated(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Widget.cls", ModuleKind: "class"}, []byte(`Public Event Changed()
Public Sub Run()
    RaiseEvent Changed
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolver := &countingEventResolver{SymbolResolver: NewResolver([]ResolverSymbol{
		{Name: "Changed", Module: "Widget", ModuleKind: "class", Kind: "event", Visibility: "Public"},
	})}
	resolved := Resolve(doc, resolver)
	if resolver.eventCalls != 1 {
		t.Fatalf("ResolveEvent calls = %d, want one", resolver.eventCalls)
	}
	if got := resolved.Procedures[0].RaiseEvents[0].Resolution.Status; got != ResolutionMatched {
		t.Fatalf("RaiseEvent status = %q, want matched", got)
	}
	if got := resolved.Procedures[0].Calls[0].Resolution.Status; got != ResolutionMatched {
		t.Fatalf("RaiseEvent call status = %q, want matched", got)
	}
}

func TestIssue590DeclarationNamesIgnoreReturnSlotCase(t *testing.T) {
	names := declarationNames(nil, []Declaration{
		{Name: "Compute", Kind: "RETURN_SLOT", Type: "Long"},
		{Name: "value", Kind: "variable", Type: "Long"},
	})
	if len(names) != 1 || names[0] != "value" {
		t.Fatalf("declaration names = %#v, want only value", names)
	}
}
