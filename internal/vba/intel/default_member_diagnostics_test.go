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
    {"name":"Excel.Range","kind":"class","source":"xlflow","confidence":"curated"}
  ],
  "global_values": {"Range":"Excel.Range"}
}`)); err != nil {
		t.Fatal(err)
	}
	return db
}
