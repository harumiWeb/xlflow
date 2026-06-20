package vbaproject

import "testing"

// TestCheckModuleSetDetectsKindMismatch is the C3 regression: checkModuleSet must
// reject a module whose model Type disagrees with the kind declared in the
// preserved PROJECT stream, not just verify name membership. Otherwise the
// rebuilt PROJECTMODULES (driven by Type) can diverge from PROJECT text.
func TestCheckModuleSetDetectsKindMismatch(t *testing.T) {
	p := &Project{
		ProjectStreamRaw: []byte("Class=Foo\r\n"),                  // PROJECT says Foo is a Class
		Modules:          []Module{{Name: "Foo", Type: ModuleStd}}, // model says standard module
	}
	if err := checkModuleSet(p); err == nil {
		t.Fatal("expected kind mismatch error, got nil")
	}
}

func TestCheckModuleSetAcceptsMatchingKind(t *testing.T) {
	p := &Project{
		ProjectStreamRaw: []byte("Module=Foo\r\n"),
		Modules:          []Module{{Name: "Foo", Type: ModuleStd}},
	}
	if err := checkModuleSet(p); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

// TestCheckModuleSetRejectsUnknownKind covers the projectKind default branch: a
// ModuleType with no PROJECT keyword must fail loudly rather than silently skip
// the kind check (a future ModuleType added without updating projectKind).
func TestCheckModuleSetRejectsUnknownKind(t *testing.T) {
	p := &Project{
		ProjectStreamRaw: []byte("Module=Foo\r\n"),
		Modules:          []Module{{Name: "Foo", Type: ModuleType(99)}},
	}
	if err := checkModuleSet(p); err == nil {
		t.Fatal("expected error for unknown ModuleType, got nil")
	}
}
