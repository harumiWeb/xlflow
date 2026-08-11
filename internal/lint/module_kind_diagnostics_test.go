package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestModuleKindDiagnosticsMatrix(t *testing.T) {
	cases := []struct {
		name string
		kind string
		src  string
		code string
		want int
	}{
		{"event standard", "standard", "Public Event Changed()\n", "VB050", 1},
		{"event class", "class", "Public Event Changed()\n", "VB050", 0},
		{"friend standard", "standard", "Friend Sub Run()\nEnd Sub\n", "VB050", 1},
		{"friend document", "document", "Friend Sub Run()\nEnd Sub\n", "VB050", 1},
		{"friend class", "class", "Friend Sub Run()\nEnd Sub\n", "VB050", 0},
		{"implements standard", "standard", "Implements IFoo\n", "VB050", 1},
		{"implements class", "class", "Implements IFoo\n", "VB050", 0},
		{"withevents standard", "standard", "Private WithEvents App As Excel.Application\n", "VB050", 1},
		{"withevents object", "class", "Private WithEvents App As Object\n", "VB050", 1},
		{"withevents scalar", "class", "Private WithEvents App As Long\n", "VB050", 1},
		{"withevents array", "class", "Private WithEvents App() As Excel.Application\n", "VB050", 1},
		{"withevents as new", "class", "Private WithEvents App As New Excel.Application\n", "VB050", 1},
		{"withevents qualified", "class", "Private WithEvents App As Excel.Application\n", "VB050", 0},
		{"withevents unresolved", "class", "Private WithEvents App As ExternalSink\n", "VB050", 0},
		{"public const object module", "class", "Public Const Flag = True\n", "VB050", 1},
		{"public array object module", "class", "Public Values() As Long\n", "VB050", 1},
		{"public fixed string object module", "class", "Public Name As String * 10\n", "VB050", 1},
		{"public type object module", "class", "Public Type State\n Value As Long\nEnd Type\n", "VB050", 1},
		{"public declare object module", "class", "Public Declare Function GetValue Lib \"x\" () As Long\n", "VB050", 1},
		{"public scalar object module", "class", "Public Value As Long\n", "VB050", 0},
		{"me standard", "standard", "Sub Run()\n Me\nEnd Sub\n", "VB051", 1},
		{"me class", "class", "Sub Run()\n Me\nEnd Sub\n", "VB051", 0},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{Config: config.Default(), ModuleKind: tc.kind}).LintSource("Module.cls", []byte(tc.src))
			if err != nil {
				t.Fatal(err)
			}
			got := 0
			for _, issue := range issues {
				if issue.Code == tc.code {
					got++
					if issue.Kind == "" {
						t.Fatalf("%s issue has no kind: %+v", tc.code, issue)
					}
				}
			}
			if got != tc.want {
				t.Fatalf("%s diagnostics = %d, want %d: %+v", tc.code, got, tc.want, issues)
			}
		})
	}
}

func TestModuleKindDiagnosticsRespectConstantConditionalBranches(t *testing.T) {
	source := "#If False Then\nPublic Event Ignored()\n#Else\nSub Run()\nEnd Sub\n#End If\n"
	issues, err := (Linter{Config: config.Default(), ModuleKind: "standard"}).LintSource("Module.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB050" {
			t.Fatalf("inactive module declaration produced VB050: %+v", issue)
		}
	}
}

func TestModuleKindDiagnosticsIgnoreCommentsAndStringsForMe(t *testing.T) {
	source := "Sub Run()\n  ' Me in a comment\n  Debug.Print \"Me\"\n  obj.Me\nEnd Sub\n"
	issues, err := (Linter{Config: config.Default(), ModuleKind: "standard"}).LintSource("Module.cls", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB051" {
			t.Fatalf("comment/string Me produced VB051: %+v", issue)
		}
	}
}
