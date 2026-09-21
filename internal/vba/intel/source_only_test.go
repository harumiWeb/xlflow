package intel

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestSourceOnlyFormControlsIgnoreHostSidecar(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "CustomerForm.frm"), []byte(`VERSION 5.00
Begin VB.UserForm CustomerForm
   Begin MSForms.TextBox hostOnly
   End
End
Attribute VB_Name = "CustomerForm"
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := Analyzer{
		RootDir:    root,
		Config:     config.Default(),
		SourceOnly: true,
	}
	codeBehind := Document{
		Path:       filepath.Join(codeDir, "CustomerForm.bas"),
		ModuleKind: "form",
		Source:     "Option Explicit\nPrivate Sub UserForm_Initialize()\n    Me.hostOnly.Text = \"source\"\nEnd Sub\n",
	}
	if got := analyzer.formControls(codeBehind); len(got) != 0 {
		t.Fatalf("source-only form controls = %+v, want no host sidecar controls", got)
	}

	virtualForm := Document{
		Path:       "virtual/forms/CustomerForm.frm",
		ModuleKind: "form",
		Source: `VERSION 5.00
Begin VB.UserForm CustomerForm
   Begin MSForms.TextBox suppliedOnly
   End
End
Attribute VB_Name = "CustomerForm"
`,
	}
	controls := analyzer.formControls(virtualForm)
	if len(controls) != 1 || !strings.EqualFold(controls[0].Name, "suppliedOnly") {
		t.Fatalf("source-only virtual form controls = %+v, want suppliedOnly", controls)
	}
}

func TestSourceOnlyWorkspaceUsesSuppliedDocumentsOnly(t *testing.T) {
	root := t.TempDir()
	modulesDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modulesDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(modulesDir, "Decoy.bas"), []byte("Public Sub HostOnly()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Src.Modules = filepath.ToSlash(filepath.Join("src", "modules"))
	open := []Document{{
		Path:       "virtual/Main.bas",
		ModuleKind: "standard",
		Source:     "Public Sub SuppliedOnly()\nEnd Sub\n",
	}}

	sourceOnly := Analyzer{RootDir: root, Config: cfg, SourceOnly: true}
	docs, err := sourceOnly.workspaceDocuments(open)
	if err != nil {
		t.Fatal(err)
	}
	if len(docs) != 1 || docs[0].Path != open[0].Path {
		t.Fatalf("source-only workspace documents = %+v, want supplied document only", docs)
	}
	symbols, err := sourceOnly.WorkspaceSymbolsContext(context.Background(), open, "")
	if err != nil {
		t.Fatal(err)
	}
	if !hasSymbolNamed(symbols, "SuppliedOnly") || hasSymbolNamed(symbols, "HostOnly") {
		t.Fatalf("source-only workspace symbols = %+v, want supplied only", symbols)
	}

	filesystemBacked := Analyzer{RootDir: root, Config: cfg}
	filesystemSymbols, err := filesystemBacked.WorkspaceSymbolsContext(context.Background(), open, "")
	if err != nil {
		t.Fatal(err)
	}
	if !hasSymbolNamed(filesystemSymbols, "HostOnly") {
		t.Fatalf("filesystem-backed workspace symbols = %+v, want HostOnly", filesystemSymbols)
	}
}

func hasSymbolNamed(symbols []Symbol, name string) bool {
	for _, symbol := range symbols {
		if strings.EqualFold(symbol.Name, name) {
			return true
		}
	}
	return false
}
