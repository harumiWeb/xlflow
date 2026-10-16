package events

import (
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestCatalogueHas222DeclarationsAcross15TypesWithOneDefaultEach(t *testing.T) {
	t.Parallel()

	declarations := Declarations()
	if len(declarations) != 222 {
		t.Fatalf("catalogue has %d declarations, want 222", len(declarations))
	}
	types := SupportedTypes()
	if len(types) != 15 {
		t.Fatalf("catalogue has %d host types, want 15: %v", len(types), types)
	}
	defaults := make(map[string]int, len(types))
	seen := make(map[string]struct{}, len(declarations))
	for _, declaration := range declarations {
		key := strings.ToLower(declaration.Type + "." + declaration.Name)
		if _, duplicate := seen[key]; duplicate {
			t.Fatalf("duplicate catalogue declaration %s", key)
		}
		seen[key] = struct{}{}
		if declaration.IsDefault {
			defaults[strings.ToLower(declaration.Type)]++
		}
	}
	for _, typeName := range types {
		if defaults[strings.ToLower(typeName)] != 1 {
			t.Errorf("%s has %d defaults, want exactly one", typeName, defaults[strings.ToLower(typeName)])
		}
	}
	if !hasDefault(declarations, "UserForm", "Click") {
		t.Error("UserForm Click is not the default")
	}
	if !hasDefault(declarations, "Image", "BeforeDragOver") {
		t.Error("Image BeforeDragOver is not the default")
	}
	if hasDeclaration(declarations, "Image", "Click") || hasDeclaration(declarations, "Image", "DblClick") {
		t.Error("hidden Image Click or DblClick was exposed")
	}
}

func TestEveryCatalogueDeclarationParsesWithoutRecovery(t *testing.T) {
	t.Parallel()

	for _, declaration := range Declarations() {
		t.Run(declaration.Type+"/"+declaration.Name, func(t *testing.T) {
			signature := strings.Replace(declaration.Signature, "{object}", declaration.Type+"1", 1)
			source := []byte(signature + "\r\nEnd Sub\r\n")
			document, err := vbaast.ParseDocument("UserFormEvents.bas", source)
			if err != nil {
				t.Fatalf("ParseDocument: %v", err)
			}
			defer document.Close()
			if err := document.Read(func(view vbaast.ParsedView) error {
				if view.HasError || view.HasMissing {
					t.Errorf("observed declaration was recovered by parser: %s", declaration.Signature)
				}
				return nil
			}); err != nil {
				t.Fatalf("ParsedDocument.Read: %v", err)
			}
		})
	}
}

func TestCataloguePreservesObservedImplicitByRefSignatures(t *testing.T) {
	t.Parallel()

	want := map[string]string{
		"UserForm.QueryClose": "Private Sub {object}_QueryClose(Cancel As Integer, CloseMode As Integer)",
		"UserForm.Zoom":       "Private Sub {object}_Zoom(Percent As Integer)",
	}
	for _, declaration := range Declarations() {
		key := declaration.Type + "." + declaration.Name
		if expected, ok := want[key]; ok && declaration.Signature != expected {
			t.Errorf("%s signature = %q, want observed spelling %q", key, declaration.Signature, expected)
		}
	}
	for key := range want {
		found := false
		for _, declaration := range Declarations() {
			if declaration.Type+"."+declaration.Name == key {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("catalogue is missing %s", key)
		}
	}
}

func hasDefault(declarations []Declaration, typeName, eventName string) bool {
	for _, declaration := range declarations {
		if strings.EqualFold(declaration.Type, typeName) && strings.EqualFold(declaration.Name, eventName) {
			return declaration.IsDefault
		}
	}
	return false
}

func hasDeclaration(declarations []Declaration, typeName, eventName string) bool {
	for _, declaration := range declarations {
		if strings.EqualFold(declaration.Type, typeName) && strings.EqualFold(declaration.Name, eventName) {
			return true
		}
	}
	return false
}
