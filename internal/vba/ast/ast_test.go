package ast

import (
	"errors"
	"testing"
)

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
