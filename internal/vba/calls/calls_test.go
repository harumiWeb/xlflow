package calls

import (
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestExtractParsedReturnsRawCallSitesAndKeepsDocumentOpen(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    BuildReport 1, 2
    Call SaveReport(Verbose:=True)
    result = CalculateTotal(1, 2)
    obj.DoSomething result
    Set item = New Customer
    CommandButton1_Click
End Sub
`)
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		t.Fatal(err)
	}

	parsedResult, err := ExtractParsed(SourceOptions{
		RootDir:    dir,
		Path:       path,
		ModuleKind: "standard",
	}, doc)
	if err != nil {
		t.Fatal(err)
	}
	legacyResult, err := extractParsedLegacy(SourceOptions{
		RootDir:    dir,
		Path:       path,
		ModuleKind: "standard",
	}, doc)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(parsedResult, legacyResult) {
		t.Fatalf("procedure IR projection changed call extraction:\nIR=%+v\nlegacy=%+v", parsedResult, legacyResult)
	}
	if parsedResult.Path != "Main.bas" || parsedResult.ModuleName != "Main" || parsedResult.ModuleKind != "standard" {
		t.Fatalf("unexpected file metadata: %+v", parsedResult)
	}
	for _, want := range []string{
		"BuildReport",
		"SaveReport",
		"CalculateTotal",
		"obj.DoSomething",
		"New Customer",
		"CommandButton1_Click",
	} {
		assertCallSite(t, parsedResult.CallSites, want)
	}
	save := assertCallSite(t, parsedResult.CallSites, "SaveReport")
	if save.Arguments.Count != 1 || len(save.Arguments.Named) != 1 ||
		save.Arguments.Named[0].Name != "Verbose" || save.Arguments.Named[0].ValueText != "True" {
		t.Fatalf("unexpected named arguments: %+v", save.Arguments)
	}
	if err := doc.Read(func(view vbaast.ParsedView) error {
		if view.Path != path || len(view.Source) == 0 || view.Root == nil {
			t.Fatalf("unexpected caller-owned document view: %+v", view)
		}
		return nil
	}); err != nil {
		t.Fatalf("ExtractParsed closed caller-owned document: %v", err)
	}

	sourceResult, err := ExtractSource(SourceOptions{
		RootDir:    dir,
		Path:       path,
		ModuleKind: "standard",
	}, source)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(parsedResult, sourceResult) {
		t.Fatalf("ExtractParsed and ExtractSource differ:\nparsed=%+v\nsource=%+v", parsedResult, sourceResult)
	}

	doc.Close()
	if parsedResult.CallSites[0].File != "Main.bas" || parsedResult.CallSites[0].Range.StartLine == 0 {
		t.Fatalf("extracted values changed after document close: %+v", parsedResult.CallSites[0])
	}
}

func TestExtractParsedPreservesDefaultRootPathCompatibility(t *testing.T) {
	path := filepath.Join(t.TempDir(), "Main.bas")
	source := []byte("Public Sub Run()\n    Call Target\nEnd Sub\n")
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	got, err := ExtractParsed(SourceOptions{Path: path}, doc)
	if err != nil {
		t.Fatal(err)
	}
	want, err := extractParsedLegacy(SourceOptions{Path: path}, doc)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("empty RootDir path compatibility changed:\nIR=%+v\nlegacy=%+v", got, want)
	}
}

func TestExtractParsedPreservesConditionalProcedureCompatibility(t *testing.T) {
	source := []byte(`#If VBA7 Then
Public Function ConditionalRun() As Long
    ConditionalRun = TargetCall(1)
End Function
#End If
`)
	doc, err := vbaast.ParseDocument("Module1.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	got, err := ExtractParsed(SourceOptions{Path: "Module1.bas"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	want, err := extractParsedLegacy(SourceOptions{Path: "Module1.bas"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("conditional procedure compatibility changed:\nIR=%+v\nlegacy=%+v", got, want)
	}
}

func TestExtractParsedPreservesParenthesizedComparisonCallCompatibility(t *testing.T) {
	source := []byte(`Public Sub Run()
    Dim actual As Long
    Dim expected As Long
    Check (actual) = expected
End Sub
`)
	doc, err := vbaast.ParseDocument("Module1.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	got, err := ExtractParsed(SourceOptions{Path: "Module1.bas"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	want, err := extractParsedLegacy(SourceOptions{Path: "Module1.bas"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("parenthesized comparison call compatibility changed:\nIR=%+v\nlegacy=%+v", got, want)
	}
}

func TestResolverCanReResolveUnchangedCallSite(t *testing.T) {
	site := CallSite{
		File:      "src/modules/Main.bas",
		Module:    "Main",
		Callee:    Callee{Text: "Target", BaseName: "Target", Member: "Target"},
		Arguments: Arguments{Named: []NamedArgument{}},
	}

	unresolved := NewResolver(nil).Resolve(site)
	if unresolved.Resolution.Status != "unresolved" {
		t.Fatalf("empty resolver status = %q, want unresolved", unresolved.Resolution.Status)
	}

	one := []symbols.Symbol{
		{Name: "Target", Module: "Zeta", Kind: "sub", File: "src/modules/Zeta.bas", StartLine: 10},
	}
	oneResolver := NewResolver(one)
	matched := oneResolver.Resolve(site)
	if matched.Resolution.Status != "matched" || len(matched.Resolution.Candidates) != 1 ||
		matched.Resolution.Candidates[0].QualifiedName != "Zeta.Target" {
		t.Fatalf("single candidate resolution = %+v", matched.Resolution)
	}
	matched.Resolution.Candidates[0].QualifiedName = "mutated"
	matchedAgain := oneResolver.Resolve(site)
	if matchedAgain.Resolution.Candidates[0].QualifiedName != "Zeta.Target" {
		t.Fatalf("resolved candidates share mutable state: %+v", matchedAgain.Resolution.Candidates)
	}

	two := append(one, symbols.Symbol{
		Name: "Target", Module: "Alpha", Kind: "function", File: "src/modules/Alpha.bas", StartLine: 5,
	})
	ambiguous := NewResolver(two).Resolve(site)
	if ambiguous.Resolution.Status != "ambiguous" || len(ambiguous.Resolution.Candidates) != 2 {
		t.Fatalf("multiple candidate resolution = %+v", ambiguous.Resolution)
	}
	if ambiguous.Resolution.Candidates[0].QualifiedName != "Alpha.Target" ||
		ambiguous.Resolution.Candidates[1].QualifiedName != "Zeta.Target" {
		t.Fatalf("candidate order = %+v", ambiguous.Resolution.Candidates)
	}
	if site.Callee.Text != "Target" {
		t.Fatalf("resolver mutated raw site: %+v", site)
	}
}

func TestResolverAdaptsToProcedureIR(t *testing.T) {
	var _ procedureir.Resolver = Resolver{}
	document := procedureir.DocumentIR{
		Procedures: []procedureir.ProcedureIR{{
			Calls: []procedureir.CallSite{{
				Caller: procedureir.ProcedureRef{
					Name: "Run", Kind: procedureir.ProcedureSub, QualifiedName: "Main.Run",
				},
				Callee: procedureir.Callee{
					Text: "Target", BaseName: "Target", Member: "Target",
				},
				Resolution: procedureir.CallResolution{Status: procedureir.ResolutionNotAttempted},
			}},
		}},
	}
	resolver := NewResolver([]symbols.Symbol{{
		Name: "Target", Module: "Helpers", Kind: "sub",
		File: "src/modules/Helpers.bas", StartLine: 4,
	}})
	resolved := procedureir.Resolve(document, resolver)
	call := resolved.Procedures[0].Calls[0]
	if call.Resolution.Status != procedureir.ResolutionMatched ||
		len(call.Resolution.Candidates) != 1 ||
		call.Resolution.Candidates[0].QualifiedName != "Helpers.Target" {
		t.Fatalf("procedure IR call resolution = %+v", call.Resolution)
	}
	if document.Procedures[0].Calls[0].Resolution.Status != procedureir.ResolutionNotAttempted {
		t.Fatalf("procedure IR resolver mutated input: %+v", document)
	}
}

func TestResolverDoesNotAliasRawCallSite(t *testing.T) {
	site := CallSite{
		Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"},
		Callee: Callee{
			Text:     "obj.Target",
			BaseName: "Target",
			Receiver: stringPointer("obj"),
			Member:   "Target",
		},
		Arguments: Arguments{
			Count: 1,
			Named: []NamedArgument{{Name: "Value", ValueText: "1"}},
		},
	}
	call := NewResolver(nil).Resolve(site)
	call.Caller.Name = "Changed"
	*call.Callee.Receiver = "changed"
	call.Arguments.Named[0].Name = "Changed"

	if site.Caller.Name != "Run" || *site.Callee.Receiver != "obj" || site.Arguments.Named[0].Name != "Value" {
		t.Fatalf("resolved call aliases raw site: %+v", site)
	}
}

func TestCloneFileResultDeepCopiesCallSites(t *testing.T) {
	original := FileResult{
		Path:       "src/modules/Main.bas",
		ModuleName: "Main",
		ModuleKind: "standard",
		CallSites: []CallSite{{
			File:   "src/modules/Main.bas",
			Module: "Main",
			Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"},
			Callee: Callee{
				Text:     "obj.Target",
				BaseName: "Target",
				Receiver: stringPointer("obj"),
				Member:   "Target",
			},
			Arguments: Arguments{
				Count: 1,
				Named: []NamedArgument{{Name: "Value", ValueText: "1"}},
			},
		}},
	}

	clone := CloneFileResult(original)
	clone.CallSites[0].Caller.Name = "Changed"
	*clone.CallSites[0].Callee.Receiver = "changed"
	clone.CallSites[0].Arguments.Named[0].Name = "Changed"
	clone.CallSites = append(clone.CallSites, CallSite{})

	if len(original.CallSites) != 1 {
		t.Fatalf("clone shares call-site slice: %+v", original.CallSites)
	}
	site := original.CallSites[0]
	if site.Caller.Name != "Run" || *site.Callee.Receiver != "obj" || site.Arguments.Named[0].Name != "Value" {
		t.Fatalf("clone shares nested call-site state: %+v", site)
	}

	empty := CloneFileResult(FileResult{CallSites: []CallSite{}})
	if empty.CallSites == nil {
		t.Fatal("clone changed non-nil empty call-sites slice to nil")
	}
	if CloneFileResult(FileResult{}).CallSites != nil {
		t.Fatal("clone changed nil call-sites slice to non-nil")
	}
}

func TestNewResolverFromSymbolsNormalizesAndDeterministicallyOrdersCandidates(t *testing.T) {
	resolver := NewResolverFromSymbols([]ResolverSymbol{
		{Name: "Target", Module: "Main", Kind: "sub", File: filepath.Join("src", "z", "Main.bas"), Line: 7},
		{Name: "Target", Module: "Main", Kind: "sub", File: filepath.Join("src", "a", "Main.bas"), Line: 9},
		{Name: "Target", Module: "Main", Kind: "function", File: filepath.Join("src", "m", "Main.bas"), Line: 5},
		{Name: "Target", Module: "Main", Kind: "sub", File: filepath.Join("src", "a", "Main.bas"), Line: 3},
		{Name: "Target", Module: "Main", Kind: "variable", File: "ignored.bas", Line: 1},
	})

	call := resolver.Resolve(CallSite{
		Callee:    Callee{Text: "Target", BaseName: "Target", Member: "Target"},
		Arguments: Arguments{Named: []NamedArgument{}},
	})
	if call.Resolution.Status != "ambiguous" {
		t.Fatalf("status = %q, want ambiguous: %+v", call.Resolution.Status, call.Resolution)
	}
	want := []Candidate{
		{QualifiedName: "Main.Target", Kind: "function", File: "src/m/Main.bas", Line: 5},
		{QualifiedName: "Main.Target", Kind: "sub", File: "src/a/Main.bas", Line: 3},
		{QualifiedName: "Main.Target", Kind: "sub", File: "src/a/Main.bas", Line: 9},
		{QualifiedName: "Main.Target", Kind: "sub", File: "src/z/Main.bas", Line: 7},
	}
	if !reflect.DeepEqual(call.Resolution.Candidates, want) {
		t.Fatalf("candidates = %+v, want %+v", call.Resolution.Candidates, want)
	}
}

func TestResolverAdapterResolvesNonProcedureProjectSymbols(t *testing.T) {
	resolver := NewResolverFromSymbols([]ResolverSymbol{{
		Name: "SharedValue", Module: "Globals", Kind: "variable",
		Visibility: "Public", File: "src/modules/Globals.bas", Line: 3,
	}})

	resolution := resolver.ResolveSymbol(procedureir.SymbolReference{
		Name: "SharedValue",
		Caller: procedureir.ProcedureRef{
			Name: "Run", Kind: procedureir.ProcedureSub, QualifiedName: "Main.Run",
		},
	})
	if resolution.Scope != procedureir.ScopeProject {
		t.Fatalf("scope = %q, want project: %+v", resolution.Scope, resolution)
	}
	want := []procedureir.Candidate{{
		QualifiedName: "Globals.SharedValue", Kind: "variable",
		File: "src/modules/Globals.bas", Line: 3,
	}}
	if !reflect.DeepEqual(resolution.Candidates, want) {
		t.Fatalf("candidates = %+v, want %+v", resolution.Candidates, want)
	}
}

func TestResolverRespectsPrivateProcedureVisibility(t *testing.T) {
	resolver := NewResolverFromSymbols([]ResolverSymbol{
		{Name: "Hidden", Module: "PrivateModule", Kind: "sub", Visibility: "Private", File: "src/modules/PrivateModule.bas", Line: 2},
		{Name: "PublicTarget", Module: "PrivateModule", Kind: "sub", Visibility: "Public", File: "src/modules/PrivateModule.bas", Line: 6},
	})

	sameModule := resolver.Resolve(CallSite{
		Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "PrivateModule.Run"},
		Callee: Callee{Text: "Hidden", BaseName: "Hidden", Member: "Hidden"},
	})
	if sameModule.Resolution.Status != "matched" {
		t.Fatalf("same-module private call = %+v", sameModule.Resolution)
	}

	crossModule := resolver.Resolve(CallSite{
		Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "Other.Run"},
		Callee: Callee{Text: "Hidden", BaseName: "Hidden", Member: "Hidden"},
	})
	if crossModule.Resolution.Status != "unresolved" {
		t.Fatalf("cross-module private call = %+v", crossModule.Resolution)
	}

	public := resolver.Resolve(CallSite{
		Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "Other.Run"},
		Callee: Callee{Text: "PublicTarget", BaseName: "PublicTarget", Member: "PublicTarget"},
	})
	if public.Resolution.Status != "matched" {
		t.Fatalf("cross-module public call = %+v", public.Resolution)
	}
}

func TestResolverRejectsReceiverlessCrossModuleClassProcedure(t *testing.T) {
	resolver := NewResolverFromSymbols([]ResolverSymbol{{
		Name: "RestoreEvents", Module: "StateHelper", ModuleKind: "class",
		Kind: "sub", Visibility: "Public", File: "src/classes/StateHelper.cls", Line: 2,
	}})

	bare := resolver.Resolve(CallSite{
		Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"},
		Callee: Callee{Text: "RestoreEvents", BaseName: "RestoreEvents", Member: "RestoreEvents"},
	})
	if bare.Resolution.Status != "unresolved" {
		t.Fatalf("receiverless cross-module class call = %+v", bare.Resolution)
	}

	receiver := "StateHelper"
	explicit := resolver.Resolve(CallSite{
		Caller: &Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"},
		Callee: Callee{Text: "StateHelper.RestoreEvents", BaseName: "RestoreEvents", Receiver: &receiver, Member: "RestoreEvents"},
	})
	if explicit.Resolution.Status != "matched" {
		t.Fatalf("explicit class receiver call = %+v", explicit.Resolution)
	}
}

func TestNewResolverPreservesLegacySymbolAdapter(t *testing.T) {
	legacy := NewResolver([]symbols.Symbol{{
		Name:      "Target",
		Module:    "Main",
		Kind:      "sub",
		File:      filepath.Join("src", "modules", "Main.bas"),
		StartLine: 12,
	}})
	neutral := NewResolverFromSymbols([]ResolverSymbol{{
		Name:   "Target",
		Module: "Main",
		Kind:   "sub",
		File:   filepath.Join("src", "modules", "Main.bas"),
		Line:   12,
	}})
	site := CallSite{Callee: Callee{Text: "Target", BaseName: "Target", Member: "Target"}}
	if !reflect.DeepEqual(legacy.Resolve(site), neutral.Resolve(site)) {
		t.Fatalf("legacy and neutral resolvers differ:\nlegacy=%+v\nneutral=%+v", legacy.Resolve(site), neutral.Resolve(site))
	}
}

func TestResolverPreservesCallClassificationPrecedence(t *testing.T) {
	projectSymbols := []symbols.Symbol{
		{Name: "Print", Module: "Main", Kind: "sub", File: "src/modules/Main.bas", StartLine: 2},
		{Name: "RunCore", Module: "App", Kind: "sub", File: "src/modules/App.bas", StartLine: 2},
		{Name: "Len", Module: "Helpers", Kind: "function", File: "src/modules/Helpers.bas", StartLine: 2},
	}
	resolver := NewResolver(projectSymbols)
	cases := []struct {
		name   string
		callee Callee
		status string
	}{
		{
			name:   "external receiver wins over bare project name",
			callee: Callee{Text: "Debug.Print", BaseName: "Print", Receiver: stringPointer("Debug"), Member: "Print"},
			status: "external",
		},
		{
			name:   "unknown receiver stays conservative",
			callee: Callee{Text: "obj.RunCore", BaseName: "RunCore", Receiver: stringPointer("obj"), Member: "RunCore"},
			status: "member_call",
		},
		{
			name:   "qualified receiver matches project procedure",
			callee: Callee{Text: "App.RunCore", BaseName: "RunCore", Receiver: stringPointer("App"), Member: "RunCore"},
			status: "matched",
		},
		{
			name:   "project procedure wins over builtin classification",
			callee: Callee{Text: "Len", BaseName: "Len", Member: "Len"},
			status: "matched",
		},
		{
			name:   "bare builtin without project symbol",
			callee: Callee{Text: "Trim", BaseName: "Trim", Member: "Trim"},
			status: "builtin_like",
		},
		{
			name:   "effectful builtin without project symbol",
			callee: Callee{Text: "Shell", BaseName: "Shell", Member: "Shell"},
			status: "builtin_like",
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := resolver.Resolve(CallSite{Callee: tc.callee, Arguments: Arguments{Named: []NamedArgument{}}})
			if got.Resolution.Status != tc.status {
				t.Fatalf("status = %q, want %q: %+v", got.Resolution.Status, tc.status, got.Resolution)
			}
		})
	}
}

func TestExtractParsedReportsRecoveryForFileWithoutCalls(t *testing.T) {
	result, err := ExtractSource(SourceOptions{Path: "Broken.bas"}, []byte(`Attribute VB_Name = "Broken"
Option Explicit
Public Sub Broken(ByVal value As String
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(result.CallSites) != 0 {
		t.Fatalf("call sites = %+v, want empty", result.CallSites)
	}
	if !result.Parse.HasError || !result.Parse.HasMissing {
		t.Fatalf("parse recovery = %+v, want error and missing", result.Parse)
	}
	if result.CallSites == nil {
		t.Fatal("empty call sites must serialize as [] rather than null")
	}
	if result.ModuleKind != "standard" {
		t.Fatalf("module kind = %q, want standard fallback", result.ModuleKind)
	}
}

func TestResolvedCallJSONRemainsFlat(t *testing.T) {
	call := NewResolver(nil).Resolve(CallSite{
		File:      "src/modules/Main.bas",
		Module:    "Main",
		Caller:    &Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"},
		Callee:    Callee{Text: "Missing", BaseName: "Missing", Member: "Missing"},
		Arguments: Arguments{Named: []NamedArgument{}},
	})
	data, err := json.Marshal(call)
	if err != nil {
		t.Fatal(err)
	}
	var got map[string]json.RawMessage
	if err := json.Unmarshal(data, &got); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"file", "module", "caller", "callee", "arguments", "range", "parse", "resolution"} {
		if _, ok := got[key]; !ok {
			t.Fatalf("resolved call JSON missing %q: %s", key, data)
		}
	}
	if _, nested := got["CallSite"]; nested {
		t.Fatalf("resolved call JSON contains nested CallSite: %s", data)
	}
	var args struct {
		Named []NamedArgument `json:"named"`
	}
	if err := json.Unmarshal(got["arguments"], &args); err != nil {
		t.Fatal(err)
	}
	if args.Named == nil {
		t.Fatalf("named arguments serialized as null: %s", data)
	}
}

func TestInspectExtractsRepresentativeCallSites(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	classDir := filepath.Join(dir, "src", "classes")
	formDir := filepath.Join(dir, "src", "forms")
	for _, path := range []string{moduleDir, classDir, formDir} {
		if err := os.MkdirAll(path, 0o755); err != nil {
			t.Fatal(err)
		}
	}
	main := `Attribute VB_Name = "Main"
Option Explicit
Public Sub RunReport()
    BuildReport 1, 2
    Call SaveReport(Verbose:=True)
    ParenthesizedCall(1, 2)
    result = CalculateTotal(1, 2)
    Debug.Print result
    obj.DoSomething result
    Application.WorksheetFunction.Sum(values)
    Set item = New Customer
    CommandButton1_Click
End Sub

Public Function CalculateTotal(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    CalculateTotal = leftValue + rightValue
End Function
`
	report := `Attribute VB_Name = "ReportBuilder"
Option Explicit
Public Sub BuildReport(ByVal first As Long, ByVal second As Long)
End Sub
Public Sub SaveReport(Optional ByVal Verbose As Boolean = False)
End Sub
Public Sub ParenthesizedCall(ByVal first As Long, ByVal second As Long)
End Sub
`
	customer := `VERSION 1.0 CLASS
Attribute VB_Name = "Customer"
Option Explicit
`
	form := `VERSION 5.00
Begin {00000000-0000-0000-0000-000000000000} UserForm1
End
Attribute VB_Name = "UserForm1"
Option Explicit
Public Sub CommandButton1_Click()
End Sub
`
	mustWrite(t, filepath.Join(moduleDir, "Main.bas"), main)
	mustWrite(t, filepath.Join(moduleDir, "ReportBuilder.bas"), report)
	mustWrite(t, filepath.Join(classDir, "Customer.cls"), customer)
	mustWrite(t, filepath.Join(formDir, "UserForm1.frm"), form)

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 4 {
		t.Fatalf("files = %d, want 4", result.Summary.Files)
	}
	assertCall(t, result.Calls, "BuildReport", "matched", 2)
	save := assertCall(t, result.Calls, "SaveReport", "matched", 1)
	if len(save.Arguments.Named) != 1 || save.Arguments.Named[0].Name != "Verbose" || save.Arguments.Named[0].ValueText != "True" {
		t.Fatalf("unexpected named arguments: %+v", save.Arguments.Named)
	}
	assertCall(t, result.Calls, "ParenthesizedCall", "matched", 2)
	assertCall(t, result.Calls, "CalculateTotal", "matched", 2)
	debug := assertCall(t, result.Calls, "Debug.Print", "external", 1)
	if debug.Callee.Receiver == nil || *debug.Callee.Receiver != "Debug" || debug.Callee.Member != "Print" {
		t.Fatalf("unexpected Debug.Print callee: %+v", debug.Callee)
	}
	assertCall(t, result.Calls, "obj.DoSomething", "member_call", 1)
	assertCall(t, result.Calls, "Application.WorksheetFunction.Sum", "external", 1)
	assertCall(t, result.Calls, "New Customer", "unresolved", 0)
	eventCall := assertCall(t, result.Calls, "CommandButton1_Click", "unresolved", 0)
	if eventCall.Caller == nil || eventCall.Caller.QualifiedName != "Main.RunReport" {
		t.Fatalf("unexpected caller: %+v", eventCall.Caller)
	}
}

func TestInspectFiltersFromAndTo(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Public Sub First()
    Target
    Other
End Sub
Public Sub Second()
    Target
End Sub
Public Sub Target()
End Sub
Public Sub Other()
End Sub
`
	mustWrite(t, filepath.Join(moduleDir, "Main.bas"), body)

	result, err := Inspect(Options{RootDir: dir, Config: cfg, From: "Main.First", To: "Target"})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Calls) != 1 {
		t.Fatalf("calls = %+v, want one filtered call", result.Calls)
	}
	if result.Calls[0].Caller == nil || result.Calls[0].Caller.Name != "First" || result.Calls[0].Callee.BaseName != "Target" {
		t.Fatalf("unexpected filtered call: %+v", result.Calls[0])
	}
}

func TestInspectClassifiesReceiverBeforeBareNameMatch(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	main := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Debug.Print "ready"
    total = Application.WorksheetFunction.Sum(values)
    App.RunCore True
End Sub

Public Sub Print()
End Sub

Public Function Sum(ByVal values As Variant) As Double
End Function
`
	app := `Attribute VB_Name = "App"
Option Explicit
Public Sub RunCore(ByVal showForm As Boolean)
End Sub
`
	mustWrite(t, filepath.Join(moduleDir, "Main.bas"), main)
	mustWrite(t, filepath.Join(moduleDir, "App.bas"), app)

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	debug := assertCall(t, result.Calls, "Debug.Print", "external", 1)
	if len(debug.Resolution.Candidates) != 0 {
		t.Fatalf("Debug.Print should not match project Print: %+v", debug.Resolution)
	}
	sum := assertCall(t, result.Calls, "Application.WorksheetFunction.Sum", "external", 1)
	if len(sum.Resolution.Candidates) != 0 {
		t.Fatalf("WorksheetFunction.Sum should not match project Sum: %+v", sum.Resolution)
	}
	runCore := assertCall(t, result.Calls, "App.RunCore", "matched", 1)
	if len(runCore.Resolution.Candidates) != 1 || runCore.Resolution.Candidates[0].QualifiedName != "App.RunCore" {
		t.Fatalf("App.RunCore should match qualified project symbol: %+v", runCore.Resolution)
	}
}

func TestInspectReportsParseRecoveryWithoutCrashing(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Broken"
Option Explicit
Public Function Broken(ByVal value As String
    Foo
End Function
`
	mustWrite(t, filepath.Join(moduleDir, "Broken.bas"), body)

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 1 || result.Summary.ParseErrors != 1 || result.Summary.MissingNodes != 1 {
		t.Fatalf("unexpected recovery summary: %+v", result.Summary)
	}
}

func TestInspectUsesConfiguredSourceDiscoveryIncludingFormSidecars(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	paths := map[string]string{
		filepath.Join(dir, cfg.Src.Modules, "Main.bas"): `Attribute VB_Name = "Main"
Public Sub Run()
    ModuleTarget
End Sub
`,
		filepath.Join(dir, cfg.Src.Classes, "Service.cls"): `VERSION 1.0 CLASS
Attribute VB_Name = "Service"
Public Sub Execute()
    ClassTarget
End Sub
`,
		filepath.Join(dir, cfg.Src.Workbook, "ThisWorkbook.cls"): `VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Private Sub Workbook_Open()
    WorkbookTarget
End Sub
`,
		filepath.Join(dir, cfg.Src.Forms, "UserForm1.frm"): `VERSION 5.00
Begin {00000000-0000-0000-0000-000000000000} UserForm1
End
Attribute VB_Name = "UserForm1"
Public Sub StaleFrmCode()
    StaleTarget
End Sub
`,
		filepath.Join(dir, cfg.Src.Forms, "code", "UserForm1.bas"): `Attribute VB_Name = "UserForm1"
Private Sub UserForm_Initialize()
    FormTarget
End Sub
`,
	}
	for path, body := range paths {
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		mustWrite(t, path, body)
	}

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 4 {
		t.Fatalf("files = %d, want standard/class/document/form sidecar: %+v", result.Summary.Files, result)
	}
	for _, callee := range []string{"ModuleTarget", "ClassTarget", "WorkbookTarget", "FormTarget"} {
		assertCall(t, result.Calls, callee, "unresolved", 0)
	}
	for _, call := range result.Calls {
		if call.Callee.Text == "StaleTarget" {
			t.Fatalf("tracked .frm code must be skipped when sidecar exists: %+v", call)
		}
	}
}

func mustWrite(t *testing.T, path, body string) {
	t.Helper()
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func assertCall(t *testing.T, calls []Call, text, status string, argCount int) Call {
	t.Helper()
	for _, call := range calls {
		if call.Callee.Text == text {
			if call.Resolution.Status != status || call.Arguments.Count != argCount {
				t.Fatalf("call %s = status %s args %d, want %s/%d: %+v", text, call.Resolution.Status, call.Arguments.Count, status, argCount, call)
			}
			if call.Range.StartLine == 0 || call.File == "" || call.Module == "" {
				t.Fatalf("call %s missing location context: %+v", text, call)
			}
			return call
		}
	}
	t.Fatalf("missing call %q in %+v", text, calls)
	return Call{}
}

func assertCallSite(t *testing.T, callSites []CallSite, text string) CallSite {
	t.Helper()
	for _, site := range callSites {
		if site.Callee.Text == text {
			if site.Range.StartLine == 0 || site.File == "" || site.Module == "" {
				t.Fatalf("call site %s missing location context: %+v", text, site)
			}
			return site
		}
	}
	t.Fatalf("missing raw call site %q in %+v", text, callSites)
	return CallSite{}
}

func stringPointer(value string) *string {
	return &value
}
