package ast

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/sourceencoding"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

func TestParserEntrypointsRejectInvalidUTF8WithTypedError(t *testing.T) {
	invalid := []byte("Sub Run()\n  value = \xff\nEnd Sub\n")
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	checked, err := parser.ParseChecked("Main.bas", invalid)
	assertSourceEncodingError(t, err)
	if checked != nil {
		t.Fatal("ParseChecked returned a tree for invalid UTF-8")
	}
	compat := parser.Parse("Main.bas", invalid)
	if compat == nil || compat.Err == nil {
		t.Fatalf("Parse compatibility result = %#v, want typed error", compat)
	}
	assertSourceEncodingError(t, compat.Err)

	path := filepath.Join(t.TempDir(), "Main.bas")
	if err := os.WriteFile(path, invalid, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err = parser.ParseFile(path)
	assertSourceEncodingError(t, err)

	_, err = ParseDocumentContext(context.Background(), "Main.bas", invalid)
	assertSourceEncodingError(t, err)
	previous, err := ParseDocument("Main.bas", []byte("Sub Run()\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	defer previous.Close()
	edits := []tree_sitter.InputEdit{{}}
	_, err = ParseDocumentIncremental("Main.bas", invalid, previous, edits)
	assertSourceEncodingError(t, err)
	_, err = ParseDocumentIncrementalIfAvailable("Main.bas", invalid, previous.result.Source, previous, edits)
	assertSourceEncodingError(t, err)
}

func assertSourceEncodingError(t *testing.T, err error) {
	t.Helper()
	var encodingErr *sourceencoding.Error
	if err == nil || !errors.As(err, &encodingErr) || encodingErr == nil || encodingErr.Status != sourceencoding.StatusInvalidUTF8 {
		t.Fatalf("error = %v, want sourceencoding.Error invalid_utf8", err)
	}
}

func TestParserParsesVBAAndReportsLocations(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	result := parser.Parse("Main.bas", []byte("Attribute VB_Name = \"Main\"\nPublic Sub Run()\nEnd Sub\n"))
	defer result.Close()

	if result.HasError || result.HasMissing {
		t.Fatalf("unexpected recovery state: error=%t missing=%t", result.HasError, result.HasMissing)
	}
	node := result.Root.NamedChild(2)
	if node == nil || node.Kind() != "sub_declaration" {
		t.Fatalf("unexpected node: %v", result.Root.ToSexp())
	}
	r := NodeRange(node)
	if r.StartLine != 2 || r.StartColumn != 1 || r.EndLine != 3 {
		t.Fatalf("unexpected range: %+v", r)
	}
}

func TestParserReportsErrorAndMissingRecovery(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	result := parser.Parse("Broken.bas", []byte("Public Function Foo(ByVal x As String\nEnd Function\n"))
	defer result.Close()

	if !result.HasError {
		t.Fatal("expected parse error")
	}
	if !result.HasMissing {
		t.Fatal("expected missing-node recovery")
	}
}

func TestParserRejectsDeclarationKeywordsAfterComma(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	for _, source := range []string{
		"Sub Test()\n  Dim x As Double, Dim i As Long\nEnd Sub\n",
		"Sub Test()\n  Dim b() As Byte, rEdIm b(10)\nEnd Sub\n",
	} {
		result := parser.Parse("Malformed.bas", []byte(source))
		reservedDeclarator := false
		Walk(result.Root, func(node *tree_sitter.Node) bool {
			if node.Kind() != "variable_declarator" {
				return true
			}
			name := node.ChildByFieldName("name")
			if name != nil && (strings.EqualFold(name.Utf8Text(result.Source), "dim") || strings.EqualFold(name.Utf8Text(result.Source), "redim")) {
				reservedDeclarator = true
			}
			return true
		})
		result.Close()
		if !result.HasError {
			t.Fatalf("expected parser recovery for %q", source)
		}
		if reservedDeclarator {
			t.Fatalf("reserved declaration keyword became a variable declarator for %q", source)
		}
	}

	for _, source := range []string{
		"Sub Test()\n  Dim x As Double, i As Long\nEnd Sub\n",
		"Sub Test()\n  Dim b() As Byte: ReDim b(10)\nEnd Sub\n",
	} {
		result := parser.Parse("Valid.bas", []byte(source))
		if result.HasError || result.HasMissing {
			t.Fatalf("valid declaration/control sequence recovered: error=%t missing=%t tree=%s", result.HasError, result.HasMissing, result.Root.ToSexp())
		}
		result.Close()
	}
}

func TestIsDeclarationKeywordRecoveryIsNarrow(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	malformed := parser.Parse("Malformed.bas", []byte("Sub Test()\n  Dim x As Double, Dim i As Long\nEnd Sub\n"))
	if !IsDeclarationKeywordRecovery(malformed.Root, malformed.Source) {
		t.Fatalf("expected declaration-keyword recovery: %s", malformed.Root.ToSexp())
	}
	malformed.Close()

	ordinary := parser.Parse("Broken.bas", []byte("Public Function Foo(ByVal x As String\nEnd Function\n"))
	if IsDeclarationKeywordRecovery(ordinary.Root, ordinary.Source) {
		t.Fatalf("ordinary parser recovery was misclassified: %s", ordinary.Root.ToSexp())
	}
	ordinary.Close()

	sameLineError := parser.Parse("Trailing.bas", []byte("Sub Test()\n  Dim x As Double, Dim i As Long: value =\nEnd Sub\n"))
	if !sameLineError.HasError && !sameLineError.HasMissing {
		t.Fatalf("expected trailing syntax to require parser recovery: %s", sameLineError.Root.ToSexp())
	}
	if IsDeclarationKeywordRecovery(sameLineError.Root, sameLineError.Source) {
		t.Fatalf("unrelated same-line recovery was misclassified: %s", sameLineError.Root.ToSexp())
	}
	sameLineError.Close()
}

func TestParserRepresentsIdentifierTypeCharacterAsBangIdentifier(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	typed := parser.Parse("Typed.bas", []byte("Sub Test()\n  Dim value!\nEnd Sub\n"))
	if typed.HasError || typed.HasMissing {
		t.Fatalf("identifier type character should parse without recovery: error=%t missing=%t tree=%s", typed.HasError, typed.HasMissing, typed.Root.ToSexp())
	}
	found := false
	Walk(typed.Root, func(node *tree_sitter.Node) bool {
		if node.Kind() != "bang_identifier" {
			return true
		}
		found = true
		if got := node.Utf8Text(typed.Source); got != "value!" {
			t.Errorf("bang_identifier text = %q, want %q", got, "value!")
		}
		if node.NamedChildCount() != 1 || node.NamedChild(0).Kind() != "identifier" {
			t.Errorf("bang_identifier children = %s, want one identifier child", node.ToSexp())
		}
		return false
	})
	typed.Close()
	if !found {
		t.Fatal("expected v0.14 bang_identifier declaration node")
	}
}

func TestParserAcceptsComparisonExpressionInCaseClause(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	result := parser.Parse("CaseNothing.bas", []byte(`Public Function IsAvailabilityRepro() As Boolean
    Dim viaWebSocket As Object

    Select Case False
        Case viaWebSocket Is Nothing
            IsAvailabilityRepro = True
        Case Else
            IsAvailabilityRepro = False
    End Select
End Function
`))
	defer result.Close()
	if result.HasError || result.HasMissing {
		t.Fatalf("comparison expression in Case clause should parse without recovery: error=%t missing=%t tree=%s", result.HasError, result.HasMissing, result.Root.ToSexp())
	}
	if !strings.Contains(result.Root.ToSexp(), "comparison_expression") {
		t.Fatalf("expected comparison_expression in Case clause: %s", result.Root.ToSexp())
	}
}

func TestIsNumericLiteralRecoveryRecognizesOptionalDefaults(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()
	recovered := 0
	for _, value := range []string{"1&", "&H10", "&O10", "1%", "1!", "1#", "1@", "1^"} {
		parsed := parser.Parse("Main.bas", []byte("Sub S(Optional value As Long = "+value+")\nEnd Sub\n"))
		if !parsed.HasError && !parsed.HasMissing {
			parsed.Close()
			continue
		}
		sexp := parsed.Root.ToSexp()
		got := IsNumericLiteralRecovery(parsed.Root, parsed.Source)
		parsed.Close()
		if !got {
			t.Fatalf("%q should be recognized as numeric-literal recovery: %s", value, sexp)
		}
		recovered++
	}
	if recovered == 0 {
		t.Fatal("expected at least one parser-recovered numeric literal form")
	}
	ordinary := parser.Parse("Main.bas", []byte("Sub S(Optional value As Long = Foo())\nEnd Sub\n"))
	if IsNumericLiteralRecovery(ordinary.Root, ordinary.Source) {
		t.Fatal("callable default must not be numeric-literal recovery")
	}
	ordinary.Close()
}

func TestParserParsesNestedInlineLoopsWithSharedNextVariables(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	result := parser.Parse("Main.bas", []byte("Public Sub Run()\nFor outer = 1 To 2: For inner = 1 To 2: Next inner, outer\nEnd Sub\n"))
	defer result.Close()

	if result.HasError || result.HasMissing {
		t.Fatalf("unexpected recovery state: error=%t missing=%t tree=%s", result.HasError, result.HasMissing, result.Root.ToSexp())
	}
}

func TestParserParsesVBEProcedureAttributesWithContextualNames(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	result := parser.Parse("Exported.cls", []byte(`Attribute VB_Name = "Exported"
Public Function Load() As String
Attribute Load.VB_Description = "Loads the document."
End Function

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name."
End Property
`))
	defer result.Close()

	if result.HasError || result.HasMissing {
		t.Fatalf("unexpected recovery state: error=%t missing=%t tree=%s", result.HasError, result.HasMissing, result.Root.ToSexp())
	}
	attributes := map[string]bool{}
	Walk(result.Root, func(node *tree_sitter.Node) bool {
		if node.Kind() != "attribute_statement" {
			return true
		}
		name := node.ChildByFieldName("name")
		if name != nil {
			attributes[name.Utf8Text(result.Source)] = name.Kind() == "qualified_member_expression"
		}
		return true
	})
	for _, want := range []string{"Load.VB_Description", "Name.VB_Description"} {
		if !attributes[want] {
			t.Fatalf("procedure Attribute %q was not parsed as a qualified member expression: %+v", want, attributes)
		}
	}
}

func TestParserParsesVBEMultiSegmentProcedureAttributes(t *testing.T) {
	parser, err := NewParser()
	if err != nil {
		t.Fatal(err)
	}
	defer parser.Close()

	result := parser.Parse("Exported.bas", []byte(`Public Sub Foo()
Attribute Foo.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub
`))
	defer result.Close()

	if result.HasError || result.HasMissing {
		t.Fatalf("unexpected recovery state: error=%t missing=%t tree=%s", result.HasError, result.HasMissing, result.Root.ToSexp())
	}
	found := false
	Walk(result.Root, func(node *tree_sitter.Node) bool {
		if node.Kind() != "attribute_statement" {
			return true
		}
		name := node.ChildByFieldName("name")
		if name == nil || name.Utf8Text(result.Source) != "Foo.VB_ProcData.VB_Invoke_Func" {
			return true
		}
		found = name.Kind() == "qualified_member_expression" && name.NamedChildCount() == 3
		return true
	})
	if !found {
		t.Fatalf("multi-segment procedure Attribute was not parsed as a qualified member expression: %s", result.Root.ToSexp())
	}
}

func TestParsedDocumentOwnsRecoveryStateAndClosesAfterReaders(t *testing.T) {
	doc, err := ParseDocument("Broken.bas", []byte("Public Function Foo(ByVal x As String\nEnd Function\n"))
	if err != nil {
		t.Fatal(err)
	}
	started := make(chan struct{})
	release := make(chan struct{})
	done := make(chan error, 1)
	go func() {
		done <- doc.Read(func(view ParsedView) error {
			if !view.HasError || !view.HasMissing || view.Root == nil || view.Path != "Broken.bas" {
				t.Errorf("view = %+v", view)
			}
			close(started)
			<-release
			return nil
		})
	}()
	<-started
	doc.Close()
	if err := doc.Read(func(ParsedView) error { return nil }); !errors.Is(err, ErrParsedDocumentClosed) {
		t.Fatalf("read after close = %v, want ErrParsedDocumentClosed", err)
	}
	if doc.result == nil {
		t.Fatal("close released a tree while a reader still owned it")
	}
	close(release)
	if err := <-done; err != nil {
		t.Fatal(err)
	}
	if doc.result != nil {
		t.Fatal("tree was not released after the final reader")
	}
	doc.Close()
}

func TestParsedDocumentReadContextCancelsWhileWaitingForTree(t *testing.T) {
	doc, err := ParseDocument("Main.bas", []byte("Sub Main()\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	started := make(chan struct{})
	release := make(chan struct{})
	holderDone := make(chan error, 1)
	go func() {
		holderDone <- doc.Read(func(ParsedView) error {
			close(started)
			<-release
			return nil
		})
	}()
	<-started

	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	visited := false
	if err := doc.ReadContext(ctx, func(ParsedView) error {
		visited = true
		return nil
	}); !errors.Is(err, context.Canceled) {
		t.Fatalf("queued read error = %v, want context.Canceled", err)
	}
	if visited {
		t.Fatal("canceled queued read invoked its callback")
	}
	close(release)
	if err := <-holderDone; err != nil {
		t.Fatal(err)
	}
}

func TestParseDocumentIncrementalClonesAndPreservesPreviousTree(t *testing.T) {
	oldSource := []byte("Sub A()\nEnd Sub\n")
	previous, err := ParseDocument("Main.bas", oldSource)
	if err != nil {
		t.Fatal(err)
	}
	defer previous.Close()
	var before string
	if err := previous.Read(func(view ParsedView) error {
		before = view.Root.ToSexp()
		return nil
	}); err != nil {
		t.Fatal(err)
	}
	edits := []tree_sitter.InputEdit{{
		StartByte: 4, OldEndByte: 5, NewEndByte: 5,
		StartPosition:  tree_sitter.Point{Column: 4},
		OldEndPosition: tree_sitter.Point{Column: 5},
		NewEndPosition: tree_sitter.Point{Column: 5},
	}}
	next, err := ParseDocumentIncremental("Main.bas", []byte("Sub B()\nEnd Sub\n"), previous, edits)
	if err != nil {
		t.Fatal(err)
	}
	defer next.Close()
	if err := previous.Read(func(view ParsedView) error {
		if got := view.Root.ToSexp(); got != before || string(view.Source) != string(oldSource) {
			t.Fatalf("previous tree changed: sexp=%q source=%q", got, view.Source)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
	if err := next.Read(func(view ParsedView) error {
		if string(view.Source) != "Sub B()\nEnd Sub\n" || view.Root == nil || view.HasError || view.HasMissing {
			t.Fatalf("incremental view = %+v", view)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestParseDocumentIncrementalRejectsClosedPreviousTree(t *testing.T) {
	previous, err := ParseDocument("Main.bas", []byte("Sub A()\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	previous.Close()
	_, err = ParseDocumentIncremental("Main.bas", []byte("Sub B()\nEnd Sub\n"), previous, []tree_sitter.InputEdit{{}})
	if !errors.Is(err, ErrIncrementalParseUnavailable) {
		t.Fatalf("incremental parse error = %v, want ErrIncrementalParseUnavailable", err)
	}
}

func TestParseDocumentIncrementalIfAvailableDoesNotWaitForReader(t *testing.T) {
	oldSource := []byte("Sub A()\nEnd Sub\n")
	previous, err := ParseDocument("Main.bas", oldSource)
	if err != nil {
		t.Fatal(err)
	}
	defer previous.Close()
	started := make(chan struct{})
	release := make(chan struct{})
	done := make(chan error, 1)
	go func() {
		done <- previous.Read(func(ParsedView) error {
			close(started)
			<-release
			return nil
		})
	}()
	<-started
	edits := []tree_sitter.InputEdit{{
		StartByte: 4, OldEndByte: 5, NewEndByte: 5,
		StartPosition: tree_sitter.Point{Column: 4}, OldEndPosition: tree_sitter.Point{Column: 5}, NewEndPosition: tree_sitter.Point{Column: 5},
	}}
	if _, err := ParseDocumentIncrementalIfAvailable("Main.bas", []byte("Sub B()\nEnd Sub\n"), oldSource, previous, edits); !errors.Is(err, ErrIncrementalParseUnavailable) {
		t.Fatalf("busy incremental parse error = %v, want ErrIncrementalParseUnavailable", err)
	}
	close(release)
	if err := <-done; err != nil {
		t.Fatal(err)
	}
}
