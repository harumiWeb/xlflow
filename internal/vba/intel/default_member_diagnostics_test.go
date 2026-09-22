package intel

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestDefaultMemberDiagnosticsClassifyKnownIndexedUnboundAndBangAccess(t *testing.T) {
	db := defaultMemberTestDB(t)
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	cfg.Analyze.DetectUnboundDefaultMemberAccess = true
	cfg.Analyze.DetectBangNotation = true
	analyzer := Analyzer{Config: cfg, DB: db}
	doc := Document{Path: "Main.bas", Source: `Option Explicit
Public Sub Run()
    Dim result As Variant
    Dim item As Widget
    Dim bag As Bag
    Dim dynamicBag As DynamicBag
    Dim late As Object
    Dim target As Widget
    result = item
    result = bag("key")
    result = dynamicBag("key")
	result = late
	result = record!Name
	result = "record!IgnoredString"
	' record!IgnoredComment
	Rem record!IgnoredRemComment
	Set target = item
    result = item.Value
    result = Range("A1")
End Sub
`}

	diagnostics, err := analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	want := map[string]int{"VBA253": 3, "VBA254": 1, "VBA255": 1}
	for _, diagnostic := range diagnostics {
		want[diagnostic.Code]--
		if diagnostic.DefaultMember == nil {
			t.Fatalf("diagnostic lacks default-member context: %+v", diagnostic)
		}
		if diagnostic.Range.End.Line < diagnostic.Range.Start.Line || diagnostic.Range.End.Character <= diagnostic.Range.Start.Character {
			t.Fatalf("diagnostic has invalid range: %+v", diagnostic)
		}
	}
	for code, remaining := range want {
		if remaining != 0 {
			t.Fatalf("%s remaining = %d; diagnostics=%+v", code, remaining, diagnostics)
		}
	}
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == "VBA253" && diagnostic.Range.Start.Line == 9 {
			if diagnostic.DefaultMember.Kind != "recursive" || diagnostic.DefaultMember.Depth != 2 {
				t.Fatalf("indexed recursive context = %+v", diagnostic.DefaultMember)
			}
		}
	}
}

func TestDefaultMemberDiagnosticsUseCompleteGeneratedAbsenceForRuntimeFailure(t *testing.T) {
	db := defaultMemberTestDB(t)
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	analyzer := Analyzer{Config: cfg, DB: db}
	doc := Document{Path: "Main.bas", Source: `Public Sub Run()
    Dim result As Variant
    Dim broken As Broken
    Set broken = New Broken
    result = broken
End Sub
`}

	diagnostics, err := analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(diagnostics) != 1 || diagnostics[0].Code != "VBA249" || diagnostics[0].DefaultMember == nil || diagnostics[0].DefaultMember.Binding != "invalid" {
		t.Fatalf("complete generated absence diagnostics = %+v", diagnostics)
	}

	uninitialized := Document{Path: "Main.bas", Source: `Public Sub Run()
    Dim result As Variant
    Dim broken As Broken
    result = broken
End Sub
`}
	diagnostics, err = analyzer.DefaultMemberDiagnosticsContext(t.Context(), uninitialized)
	if err != nil {
		t.Fatal(err)
	}
	if len(diagnostics) != 0 {
		t.Fatalf("uninitialized receiver must remain owned by VBA202: %+v", diagnostics)
	}

	analyzer.TypeDBResolutionIncomplete = true
	diagnostics, err = analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(diagnostics) != 0 {
		t.Fatalf("incomplete TypeDB must fail open: %+v", diagnostics)
	}
}

func TestDefaultMemberDiagnosticsReportInitializedDefaultMemberCycle(t *testing.T) {
	db := defaultMemberTestDB(t)
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	analyzer := Analyzer{Config: cfg, DB: db}
	doc := Document{Path: "Main.bas", Source: `Public Sub Run()
    Dim result As Variant
    Dim cycle As Cycle
    Set cycle = New Cycle
    result = cycle
End Sub
`}
	diagnostics, err := analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(diagnostics) != 1 || diagnostics[0].Code != "VBA249" || diagnostics[0].DefaultMember == nil || diagnostics[0].DefaultMember.Kind != "recursive" {
		t.Fatalf("default-member cycle diagnostics = %+v", diagnostics)
	}
}

func TestDefaultMemberDiagnosticsTreatNamedEnumResultAsValue(t *testing.T) {
	db := defaultMemberTestDB(t)
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	analyzer := Analyzer{Config: cfg, DB: db}
	doc := Document{Path: "Main.bas", Source: `Public Sub Run()
	Dim result As Variant
	Dim choice As EnumChoice
	Set choice = New EnumChoice
	result = choice
End Sub
`}

	diagnostics, err := analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(diagnostics) != 1 || diagnostics[0].Code != "VBA253" || diagnostics[0].DefaultMember == nil || diagnostics[0].DefaultMember.Member != "Value" {
		t.Fatalf("named enum result diagnostics = %+v", diagnostics)
	}
}

func TestDefaultMemberDiagnosticsValidateIndexedDefaultMemberArity(t *testing.T) {
	db := defaultMemberTestDB(t)
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	analyzer := Analyzer{Config: cfg, DB: db}
	doc := Document{Path: "Main.bas", Source: `Public Sub Run()
	Dim result As Variant
	Dim bag As Bag
	Set bag = New Bag
	result = bag
	result = bag()
	result = bag("key")
	result = bag("key", "extra")
End Sub
`}

	diagnostics, err := analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	want := map[int]string{6: "VBA253"}
	for _, diagnostic := range diagnostics {
		line := diagnostic.Range.Start.Line
		if code, ok := want[line]; !ok || code != diagnostic.Code {
			t.Fatalf("unexpected arity diagnostic = %+v; want=%v", diagnostic, want)
		}
		delete(want, line)
	}
	if len(want) != 0 {
		t.Fatalf("missing arity diagnostics: %v; got=%+v", want, diagnostics)
	}
}

func TestDefaultMemberArgumentCompatibility(t *testing.T) {
	tests := []struct {
		name           string
		parameters     []vbadb.ParamInfo
		argumentCounts map[int]bool
	}{
		{
			name:           "required and optional",
			parameters:     []vbadb.ParamInfo{{Name: "Key"}, {Name: "Fallback", Optional: true}},
			argumentCounts: map[int]bool{0: false, 1: true, 2: true, 3: false},
		},
		{
			name:           "param array",
			parameters:     []vbadb.ParamInfo{{Name: "Keys", ParamArray: true}},
			argumentCounts: map[int]bool{0: true, 1: true, 3: true},
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			member := vbadb.MemberInfo{Name: "Item", Parameters: test.parameters}
			for argumentCount, want := range test.argumentCounts {
				if got := defaultMemberAcceptsArguments(member, argumentCount); got != want {
					t.Fatalf("defaultMemberAcceptsArguments(%d) = %t, want %t", argumentCount, got, want)
				}
			}
		})
	}
}

func TestDefaultMemberDiagnosticsReportInitializedObjectUsedAsCallableWithoutDefault(t *testing.T) {
	db := defaultMemberTestDB(t)
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	analyzer := Analyzer{Config: cfg, DB: db}
	doc := Document{Path: "Main.bas", Source: `Public Sub Run()
	Dim result As Variant
	Dim broken As Broken
	Set broken = New Broken
	result = broken(1)
End Sub
`}

	diagnostics, err := analyzer.DefaultMemberDiagnosticsContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(diagnostics) != 1 || diagnostics[0].Code != "VBA249" || diagnostics[0].DefaultMember == nil || diagnostics[0].DefaultMember.Kind != "indexed" {
		t.Fatalf("callable-object diagnostics = %+v", diagnostics)
	}
}

func defaultMemberTestDB(t *testing.T) *vbadb.DB {
	t.Helper()
	db := vbadb.New()
	if err := db.MergeJSON([]byte(`{
  "types": [
    {"name":"Widget","kind":"class","source":"typelib","confidence":"generated","default_member":"Value","default_member_type":"String","properties":[{"name":"Value","return_type":"String","default":true}]},
    {"name":"Bag","kind":"class","source":"typelib","confidence":"generated","default_member":"Item","default_member_type":"Widget","properties":[{"name":"Item","return_type":"Widget","default":true,"parameters":[{"name":"Key","type":"Variant"}]}]},
    {"name":"DynamicBag","kind":"class","source":"typelib","confidence":"generated","default_member":"Item","default_member_type":"Variant","properties":[{"name":"Item","return_type":"Variant","default":true,"parameters":[{"name":"Key","type":"Variant"}]}]},
    {"name":"Broken","kind":"class","source":"typelib","confidence":"generated"},
    {"name":"Cycle","kind":"class","source":"typelib","confidence":"generated","default_member":"Self","default_member_type":"Cycle","properties":[{"name":"Self","return_type":"Cycle","default":true}]},
    {"name":"EnumChoice","kind":"class","source":"typelib","confidence":"generated","default_member":"Value","default_member_type":"Example.Choice","properties":[{"name":"Value","return_type":"Example.Choice","default":true}]},
    {"name":"Example.Choice","kind":"enum","source":"typelib","confidence":"generated"},
    {"name":"Excel.Range","kind":"class","source":"xlflow","confidence":"curated"}
  ],
  "global_values": {"Range":"Excel.Range"}
}`)); err != nil {
		t.Fatal(err)
	}
	return db
}
