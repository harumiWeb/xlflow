package coordination

import (
	"path/filepath"
	"testing"
)

func TestSamePathMatchesEquivalentMissingPaths(t *testing.T) {
	base := t.TempDir()
	direct := filepath.Join(base, "project", "book.xlsm")
	withSegments := filepath.Join(base, "project", "nested", "..", "book.xlsm")
	if !SamePath(direct, withSegments) {
		t.Fatalf("SamePath(%q, %q) = false", direct, withSegments)
	}
}

func TestSamePathDistinguishesPaths(t *testing.T) {
	base := t.TempDir()
	if SamePath(filepath.Join(base, "first.xlsm"), filepath.Join(base, "second.xlsm")) {
		t.Fatal("SamePath() matched distinct paths")
	}
}
