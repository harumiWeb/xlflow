package lint

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestProcedureLocalReferenceIndexScansProcedureOnce(t *testing.T) {
	const localCount = 128

	var source strings.Builder
	source.WriteString("Public Sub Run()\n")
	allSymbols := []symbols.Symbol{{
		Name:      "Run",
		Kind:      "sub",
		StartLine: 1,
		EndLine:   localCount + 4,
	}}
	for i := range localCount {
		line := i + 2
		fmt.Fprintf(&source, "  Dim value%d As Long\n", i)
		allSymbols = append(allSymbols, symbols.Symbol{
			Name:      fmt.Sprintf("value%d", i),
			Kind:      "local_variable",
			Parent:    "Run",
			StartLine: line,
			EndLine:   line,
		})
	}
	source.WriteString("  value0 = 1\n")
	source.WriteString("  Debug.Print value1\n")
	source.WriteString("End Sub\n")

	var stats unusedLocalReferenceIndexStats
	referenced, err := buildProcedureLocalReferenceIndexContext(context.Background(), source.String(), allSymbols, &stats)
	if err != nil {
		t.Fatal(err)
	}
	if stats.SourceNormalizations != 1 {
		t.Fatalf("source normalized %d times, want 1", stats.SourceNormalizations)
	}
	if stats.ProcedureScans != 1 {
		t.Fatalf("procedure scanned %d times for %d locals, want 1", stats.ProcedureScans, localCount)
	}
	if referenced[1] {
		t.Fatal("write-only value0 should not be indexed as a read")
	}
	if !referenced[2] {
		t.Fatal("Debug.Print value1 should be indexed as a read")
	}
	for i := 3; i < len(referenced); i++ {
		if referenced[i] {
			t.Fatalf("declaration-only %s should not be indexed as a read", allSymbols[i].Name)
		}
	}
}

func TestProcedureLocalReferenceIndexKeepsAccessorsSeparate(t *testing.T) {
	source := `Property Get Value() As Long
  Dim result As Long
  Debug.Print result
End Property
Property Let Value(ByVal rhs As Long)
  Dim result As Long
  result = rhs
End Property
`
	allSymbols := []symbols.Symbol{
		{Name: "Value", Kind: "property_get", StartLine: 1, EndLine: 4},
		{Name: "result", Kind: "local_variable", Parent: "Value", StartLine: 2, EndLine: 2},
		{Name: "Value", Kind: "property_let", StartLine: 5, EndLine: 8},
		{Name: "result", Kind: "local_variable", Parent: "Value", StartLine: 6, EndLine: 6},
	}

	referenced, err := buildProcedureLocalReferenceIndexContext(context.Background(), source, allSymbols, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !referenced[1] {
		t.Fatal("Property Get local read was not indexed")
	}
	if referenced[3] {
		t.Fatal("Property Let write-only local was incorrectly indexed from the getter")
	}
}

func TestProcedureLocalReferenceIndexSupportsUnicodeNames(t *testing.T) {
	source := "Public Sub Run()\n  Dim 合計 As Long\n  Debug.Print 合計\nEnd Sub\n"
	allSymbols := []symbols.Symbol{
		{Name: "Run", Kind: "sub", StartLine: 1, EndLine: 4},
		{Name: "合計", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
	}

	referenced, err := buildProcedureLocalReferenceIndexContext(context.Background(), source, allSymbols, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !referenced[1] {
		t.Fatal("Unicode local read was not indexed")
	}
}

func TestProcedureLocalReferenceIndexMatchesIdentifierTypeCharacters(t *testing.T) {
	source := `Public Sub Run()
  Dim text$, whole%, longValue&, singleValue!, doubleValue#, money@, longLong^
  Debug.Print text$, whole%, longValue&, singleValue!, doubleValue#, money@, longLong^
End Sub
`
	allSymbols := []symbols.Symbol{
		{Name: "Run", Kind: "sub", StartLine: 1, EndLine: 4},
		{Name: "text$", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
		{Name: "whole%", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
		{Name: "longValue&", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
		{Name: "singleValue!", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
		{Name: "doubleValue#", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
		{Name: "money@", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
		{Name: "longLong^", Kind: "local_variable", Parent: "Run", StartLine: 2, EndLine: 2},
	}

	referenced, err := buildProcedureLocalReferenceIndexContext(context.Background(), source, allSymbols, nil)
	if err != nil {
		t.Fatal(err)
	}
	for i := 1; i < len(referenced); i++ {
		if !referenced[i] {
			t.Fatalf("type-character local %s was not indexed", allSymbols[i].Name)
		}
	}
}
